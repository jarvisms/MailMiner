import base64
import re

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
            # Iterate over parts of a multipart message
            for p in range(len(msgstruct[b"BODYSTRUCTURE"][0])):
                # Get the content-disposition field if there is one.
                # Expression OR (None,) will yield the tuple if None
                # so the [0] indexing works later
                cd = (msgstruct[b"BODYSTRUCTURE"][0][p][-3] or (None,))
                # find where content-disposition says its an attachment
                if ( cd[0] == b"attachment" ):
                    # the -3rd field is a "parameter parenthesized list"
                    # for the attachment. Take the attribute/value tuple
                    # and convert it into a dict first.
                    properties = dict(
                        zip(
                            *[iter(
                                msgstruct[b"BODYSTRUCTURE"][0][p][-3][1],
                            )]*2
                        )
                    )
                    # Check if the filename  matches the regex criteria
                    regexmatch = settings["regex"].fullmatch(
                        properties[b"filename"],
                    )
                    print(
                        "File: {filename} in email ID {id} tested as {match}".format(
                            filename = properties[b"filename"].decode(),
                            id = uid,
                            match = True if regexmatch else False,
                        )
                    )
                    properties.update(
                        {
                            b"uid":uid,
                            b"regexmatch":regexmatch,
                            b"encoding":msgstruct[b"BODYSTRUCTURE"][0][p][5],
                            b"textsize":msgstruct[b"BODYSTRUCTURE"][0][p][6],
                            b"part":p+1,
                        }
                        )
                    # If encoding is correct and filename matches the regex
                    if properties[b"encoding"] == b"base64" and regexmatch:
                        # It should already have a dict so
                        # add this filename and email body part number
                        # (which start from 1)
                        filedetails[uid] = properties
    return filedetails

def FetchAttachments(server,filedetails,filedata=[]):
    """Given a server and list of email parts,
    Returns a list of dicts containing of real file data contents
    This function is memory inefficient"""
    # Filedata will continue to grow with successive calls to
    # FetchAttachments unless filedata is explicitly parsed
    byparts={}
    batch={}
    for id in filedetails:
        part = filedetails[id][b"part"]
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
            print(
                "Decoded '{}'".format(
                    filedetails[uid][b"filename"].decode()
                )
            )
    return filedata


def FetchDecode(server,detail):
    """Fetches an email attachment and returned the decoded contents"""
    print(
        "Downloading email ID: {}, part: {}".format(
            detail[b"uid"],detail[b"part"],
        )
    )
    # Fetches and decodes payload immediately
    data = base64.b64decode(
        server.fetch(
            detail[b"uid"],
            ["BODY[{}]".format(detail[b"part"]).encode()],
        )[detail[b"uid"]]["BODY[{}]".format(detail[b"part"]).encode()]
    )
    print(
        "Attachment: '{}', {} bytes.".format(
            detail[b"filename"].decode(),len(data)
        )
    )
    return data

if __name__ == "__main__":
    import configparser
    import imaplib
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
            ).encode()
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
        print("Found {} attachments".format(len(filedetails)))
        # Generator Expression where each element is a copy of
        # the original filedetails dictionary along with the
        # bytedata after fetching, downloading and decoding it
        filedata = (
            {
                "bytedata":FetchDecode(
                    server,
                    filedetails[uid],
                ),
                **filedetails[uid],
            } for uid in filedetails
        )
        # Runs the named function directly from the local scope
        locals()[
            config.get(
                section,
                "converter",
                fallback="Shelve",
            )
        ](filedata,settings)
    print(server.logout().decode())

