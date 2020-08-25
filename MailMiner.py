import base64
import re
import quopri

encoded_word_regex = re.compile(r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='.encode())

# https://stackoverflow.com/questions/12739563/parsing-email-bodystructure-in-python?noredirect=1
# ----- Parsing BODYSTRUCTURE into parts dictionaries ----- #

def tuple2dict(pairs):
    """get dict from (key, value, key, value, ...) tuple"""
    if not pairs:
        return None
    return dict([(k, tuple2dict(v) if isinstance(v, tuple) else v)
                 for k, v in zip(pairs[::2], pairs[1::2])])

def parse_singlepart(var, part_no):
    """convert non-multipart into dic"""
    # Basic fields for non-multipart (Required)
    part = dict(zip(['maintype', 'subtype', 'params', 'id', 'description', 'encoding', 'size'], var[:7]), part_no=part_no)
    part['params'] = tuple2dict(part['params'])
    # Type specific fields (Required for 'message' or 'text' type)
    index = 7
    if part['maintype'].lower() in (b'message','message') and part['subtype'].lower() in (b'rfc822','rfc822'):
        part.update(zip(['envelope', 'bodystructure', 'lines'], var[7:10]))
        index = 10
    elif part['maintype'].lower() in (b'text','text'):
        part['lines'] = var[7]
        index = 8
    # Extension fields for non-multipart (Optional)
    part.update(zip(['md5', 'disposition', 'language', 'location'], var[index:]))
    part['disposition'] = tuple2dict(part['disposition'])
    return part

def parse_multipart(var, part_no):
    """convert the multipart into dict"""
    part = { 'child_parts': [], 'part_no': part_no }
    # First parse the child parts
    index = 0
    if isinstance(var[0], list):
        part['child_parts'] = [parse_part(v, ('%s.%d' % (part_no, i+1)).replace('TEXT.', '')) for i, v in enumerate(var[0])]
        index = 1
    elif isinstance(var[0], tuple):
        while isinstance(var[index], tuple):
            part['child_parts'].append(parse_part(var[index], ('%s.%d' % (part_no, index+1)).replace('TEXT.', '')))
            index += 1
    # Then parse the required field subtype and optional extension fields
    part.update(zip(['subtype', 'params', 'disposition', 'language', 'location'], var[index:]))
    part['params'] = tuple2dict(part['params'])
    part['disposition'] = tuple2dict(part['disposition'])
    return part

def parse_part(var, part_no=None):
    """Parse IMAP email BODYSTRUCTURE into nested dictionary

    See http://tools.ietf.org/html/rfc3501#section-6.4.5 for structure of email messages
    See http://tools.ietf.org/html/rfc3501#section-7.4.2 for specification of BODYSTRUCTURE
    """
    if isinstance(var[0], (tuple, list)):
        return parse_multipart(var, part_no or 'TEXT')
    else:
        return parse_singlepart(var, part_no or '1')

# ----- End of Parsing BODYSTRUCTURE into parts dictionaries ----- #

def FlatParts(parts, flat=None):
    if flat is None:
        flat = {}
    child_parts = parts.get('child_parts',[])
    parent_part = {k:v for k,v in parts.items() if k != 'child_parts'}
    flat.update({ parent_part['part_no'] : parent_part })
    for part in child_parts:
        FlatParts(part,flat)
    return flat


def DecodeFilename(filename):
    """Decodes a bytes filename to a standard string if it is
    base64 or quopri word encoded or assumed utf-8 bytes"""
    encoded = encoded_word_regex.fullmatch(filename)
    if encoded:
        charset, encoding, encoded_text = encoded.groups()
        if encoding == b'B':
            filename = base64.b64decode(encoded_text).decode(charset.decode())
        elif encoding == b'Q':
            filename = quopri.decodestring(encoded_text).decode(charset.decode())
    else:
        filename = filename.decode()
    return filename


def FindAttachments(server,settings):
    """Given a server and settings containing the criteria,
    Return a list of dicts of emails and all the various properties"""
    # Prepare to create a dictionary
    filedetails={}
    # Drop readonly to allow things to be flagged as Seen
    server.select_folder(settings["folder"], readonly=settings["readonly"])
    # Eventually, search for emails matching the various criteria
    msguids = server.search(settings["search"])
    # Get structure without downloading message
    allmsgstructs = server.fetch(msguids, ["BODYSTRUCTURE"])
    for uid in msguids:
        msgstruct = allmsgstructs[uid]
        if (b"BODYSTRUCTURE" in msgstruct):
            # Get all nested parts in a flat dictionary
            # Find parts which are of attachment disposition
            parts = FlatParts(parse_part(msgstruct[b"BODYSTRUCTURE"]))
            for p in parts:
                properties = {} # Specific for the part
                part = parts[p]
                disposition = part.get("disposition")
                if disposition is not None and b"attachment" in disposition:
                    properties = disposition
                    # Decode file name into standard string
                    properties[b"filename"] = DecodeFilename(disposition[b"attachment"][b"filename"])
                    # Check if the filename  matches the regex criteria
                    regexmatch = settings["regex"].fullmatch(properties[b"filename"])
                    print(f"File: {properties[b'filename']} in email {uid}/{p} tested as {True if regexmatch else False}")
                    properties.update(
                        {
                            b"uid":uid,
                            b"regexmatch":regexmatch,
                            b"encoding":part["encoding"],
                            b"textsize":part["size"],
                            b"part":p,
                        }
                    )
                    # If encoding is correct and filename matches the regex
                    if properties[b"encoding"] == b"base64" and regexmatch:
                        # It should already have a dict so
                        # add this filename and email body part number
                        # (which start from 1)
                        if uid in filedetails: filedetails[uid][p] = properties
                        else: filedetails[uid] = {p:properties}
    return filedetails


def FetchAttachments(server,filedetails,filedata=None):
    """Given a server and list of email parts,
    Returns a list of dicts containing of real file data contents
    This function is memory inefficient"""
    # Filedata will continue to grow with successive calls to
    # FetchAttachments unless filedata is explicitly parsed
    if filedata is None: filedata = []
    byparts={}
    batch={}
    for id in filedetails:
        for part in filedetails[id].keys():
            # For each body[part], list the UIDs for a bulk download
            byparts[part] = [id] if part not in byparts else byparts[part]+[id]
    for part in byparts:
        print("Starting to download a batch from the IMAP server...")
        # Get the actual payload remembering that IMAP Body parts are
        # indexed from 1 which has already been added,
        # this will be base64 encoded
        batch.update(
            server.fetch(
                byparts[part],
                [f"BODY[{part}]".encode()],
            )
        )
        print("Batch downloaded")
        for uid in byparts[part]:
            # Decode the data into proper bytes in each
            # batch of parts and add to the big list
            filedata.append(
                {
                    "bytedata" : base64.b64decode(
                        batch[uid][f"BODY[{part}]".encode()]
                    ),
                    **filedetails[uid]
                }
            )
            print(f"Decoded '{filedetails[uid][part][b'filename']}' from email {uid}/{part}")
    return filedata


def FetchDecode(server,detail):
    """Fetches an email attachment and returned the decoded contents"""
    print(f"Downloading email ID: {detail[b'uid']}, part: {detail[b'part']}")
    # Fetches and decodes payload immediately
    data = base64.b64decode(
        server.fetch(
            detail[b"uid"],
            [f"BODY[{detail[b'part']}]".encode()],
        )[detail[b"uid"]][f"BODY[{detail[b'part']}]".encode()]
    )
    print(f"Attachment: '{detail[b'filename']}', {len(data)} bytes.")
    return data

if __name__ == "__main__":
    import configparser
    import imaplib
    import Converters # Secondary Library where Converters functinos are defined
    from imapclient import IMAPClient
    serverconf, config = configparser.ConfigParser(), configparser.ConfigParser()
    serverconf.read(r"server.cfg")
    config.read(r"config.cfg")
    
    try:
        ServerSettings = dict(serverconf.items("Settings"))
        imapserver = ServerSettings["server"]
        username = ServerSettings["username"]
        password = ServerSettings["password"]
        tasks = [task.strip() for task in ServerSettings["tasks"].split(",")]
        if not set(tasks) <= set(config.sections()):
            raise KeyError
    except KeyError:
        print(
            """Error with server settings in config file.
Check the "Settings" section contains the "server",
"username", "password", and "tasks" options"""
        )
    
    # Use UIDs so numbers are permanent
    server = IMAPClient(imapserver, use_uid=True)
    # Standard TLS version doesnt work so make the connection by hand
    server._imap = imaplib.IMAP4_SSL(host=imapserver)
        # Actually login and print the output while at it
    print(server.login(username,password).decode())

    for section in tasks:
        settings = dict(config.items(section))
        # Subfolder of INBOX, i.e. folder = "INBOX/Siemens Energy"
        settings["folder"] = "INBOX/" + config.get(
            section,"folder",fallback="",
        )
        # Treat the folder as readonly so nothing gets marked as read/seen
        settings["readonly"] = config.getboolean(
            section,"readonly",fallback=True,
        )
        # Split words by spaces into a list, i.e. searchcriteria = ["ALL"]
        settings["search"] = config.get(
            section,"search",fallback="ALL",
        ).split(" ")
        # Convert regex string to bytes
        # i.e. regex = re.compile(rb"UVW_[0-9]{6}_to_[0-9]{6}_produced_at_[0-9]{6}\.csv")
        settings["regex"] = re.compile(
            config.get(
                section,"filename",fallback=None,
            )
        )
        if "filename" not in settings or settings["filename"] == "":
            print(
                """There was no "filename" option given, or it was empty.
A regex was expected. This section will be skipped"""
            )
            continue
        # The output file name
        settings["outfile"] = config.get(
            section,"outfile",fallback="output.csv"
        )
        print(
            f"""Section: "{section}"
    Folder: "{settings["folder"]}"
    Read Only: "{settings["readonly"]}"
    Criteria: "{settings["search"]}"
    Regex: "{settings["filename"]}"
    outfile: "{settings["outfile"]}"
    """
        )
        # Get all the attachment details
        filedetails = FindAttachments(server,settings)
        print(f"Found {len(filedetails)} attachments")
        # Generator Expression where each element is a copy of
        # the original filedetails dictionary along with the
        # bytedata after fetching, downloading and decoding it
        filedata = (
            {
                "bytedata":FetchDecode(server,part),
                **part,
            } for msg in filedetails.values() for part in msg.values()
        )
        # Runs the named function directly from the local scope
        getattr(Converters,
            config.get(
                section,
                "converter",
                fallback="Shelve",
            )
        )(filedata,settings)
    print(server.logout().decode())

