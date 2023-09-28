def ConvertNum(string):
    """Converts strings to integers preferentially,
    unless it must be represented by a float,
    otherwise None if its not a valid number"""
    try:
        f = float(string.strip())
        i = int(f)
        if f == i:
            return i
        else:
            return f
    except ValueError:
        return None

def Concatenate(filedata,settings):
    """Writes raw output to a single file from a list or generator
    yielding raw decoded files, result is just a concatenation
        
    Expects to be given an iterable giving dictionaries
    with a filename and raw bytes filedata"""
    with open(settings["outfile"], "a+b") as output:
        # Iterates through generator which will fetch and decode each item
        for file in filedata:
            print(f"Processing file '{file[b'filename']}'")
            output.write(file["bytedata"])    # Actually writes the data
    return None

def MetOfficeWeather(filedata,settings):
    """Write output to a csv file of predefined format from a concatenation
    of multiple file attachements from the MetOffice Weather Station csv format
    
    Expects to be given an iterable giving dictionaries
    with a filename and raw bytes filedata"""
    import csv
    import re
    from datetime import datetime
    from operator import itemgetter
    # Precompile regex to capture text before line ends
    regexsplitlines = re.compile(b'^(.+?)(?:\r\n|\r|\n|$)+', flags=re.MULTILINE)
    with open(settings["outfile"], "a+", newline="") as output:
        NeedHeaders = not(output.tell()) # in append mode, tell==0 if new file
        # Create the Output CSV File
        csvout = csv.writer(output, dialect="excel")
        # Only if needed, apply the Headings
        # skipping the first 2 coloumns as these are discarded
        if NeedHeaders:
            csvout.writerow(settings["headers"].split(","))
        r=0
        for file in filedata:
            try:
                print(f"Processing file '{file[b'filename']}'")
                # Take the raw file which is just bytes,
                # chop it into lines and feed that into csv module.
                # regex finditer is more memory efficient than str.splitlines()
                # which does not scale to large file sizes.
                csvinput = csv.reader(
                    m.group(1).decode() for m in regexsplitlines.finditer(
                        file["bytedata"],
                    )
                )
                for line in csvinput:
                    # Look for the phrase followed by Date on the next line
                    if (line[0] == "Hourly Summary Data"
                        and csvinput.__next__()[0] == "Date"):
                            # Break out of for loop to start data gathering
                            break
                # Should now be on data
                for line in csvinput:
                    # Assumes dates are valid %d/%m/%Y,%H%M and will become
                    # %d/%m/%Y %H:%M with leading zeros inserted if missing.
                    # All numbers are converted or left blank
                    # but only the non blank coloumns are chosen
                    csvout.writerow(
                        [datetime.strptime(
                            " ".join(line[:2]), "%d/%m/%Y %H%M"
                                ).strftime("%d/%m/%Y %H:%M")]
                        + [ConvertNum(item) for item in itemgetter(
                            4,5,6,7,8,9,10,11,13,14)(line)])
                    r+=1
                print(f"{r} unique rows written so far")
            except Exception as e:
                print(f"""Encountered some issue with '{file[b"filename"]}',
but {r} rows written so far.
Error: {e}""")
    print(f"Finished. {r} unique rows written\n")
    return None

def Bablake(filedata,settings):
    """Write output to a csv file of predefined format from a concatenation
    of multiple file attachements from the Bablake Weather Station Excel format
    
    Expects to be given an iterable giving dictionaries
    with a filename and raw bytes filedata and a regex match
    
    The Regex for the filename definition
    must label the month and year parts"""
    import csv
    from datetime import datetime, timedelta
    from xlrd import open_workbook
    headers = settings["headers"].split(",")
    ncols = len(headers)+2
    r=0
    seen = set()
    with open(settings["outfile"], "a+", newline="") as output:
        NeedHeaders = not(output.tell()) # in append mode, tell==0 if new file
        # Create the Output CSV File
        csvout = csv.writer(output, dialect="excel")
        # Only if needed, apply the Headings
        # skipping the first 2 coloumns as these are discarded
        if NeedHeaders:
            csvout.writerow(settings["headers"].split(","))
        for file in filedata:
            try:
                filename = file[b"filename"]
                print(f"Processing file '{filename}'")
                wb = open_workbook(
                    filename=filename,
                    file_contents=file["bytedata"],
                )
                # Open the next input workbook
                s = wb.sheet_by_index(0)
                # Get the month and year that file attachment
                # is intended for based on the filename.
                # Start by getting a byted parts dictionary
                # of the date parts from the filename
                dateparts = file[b"regexmatch"].groupdict()
                # Convert all values to normal strings
                dateparts = {
                    part:dateparts[part] for part in dateparts
                }
                # Expand the dictionary to a string and  convert to a date
                filedate = datetime.strptime(
                    "{month} {year}".format(**dateparts),"%B %Y"
                )
                # Work out the 1st Jan of that year
                baseyear = datetime(filedate.year,1,1)
                # Work through all rows except the heading unless it's been
                # seen already reformat the date and time, and write to csv file
                for row in range(1,s.nrows):
                    # Store the values as a list
                    rowvals = s.row_values(row)
                    # Only proceed if we havent seen this line before and
                    # for "1" data, anything else is averages/totals etc.
                    if rowvals[0] == 1 and tuple(rowvals) not in seen:
                        # Keep track of what's been seen
                        seen.add(tuple(rowvals))
                        # No of days and hours rom new year
                        offset = timedelta(
                            days = (rowvals[1])-1,
                            hours = (rowvals[2]/100.0)-1,
                        )
                        # Add offset to new year to get actual date and time
                        dt = baseyear + offset
                        if ( filedate.month == 1 and dt.month == 12 ):
                            # If file is for January but data for December,
                            # data is actually for the previous year
                            dt.replace(year=baseyear.year-1)
                        elif ( filedate.month == 12 and dt.month == 1 ):
                            # If file is for December but data for January,
                            # data is actually for the next year
                            dt.replace(year=baseyear.year+1)
                        # Dump it to the file
                        csvout.writerow(
                            [dt.strftime("%d/%m/%Y %H:%M")]
                            + s.row_values(row,3,ncols)
                        )
                        r+=1
                print(f"{r} unique rows written so far")
            except Exception as e:
                print(f"""Encountered some issue with '{file[b"filename"]}', but {r} rows written so far.
Error: {e}""")
            # Clean up and state number of rows written
            finally:
                wb.release_resources()
    print(f"Finished. {r} unique rows written\n")
    return None

def Shelve(filedata,settings):
    """Writes raw output to a shelve file
        
    Expects to be given an iterable giving dictionaries
    with a filename and raw bytes filedata"""
    import shelve
    # Open shelf read/write or create
    myshelf=shelve.open(settings["outfile"],"c")
    myshelf["settings"] = settings
    myshelf["filedata"] = []
    for file in filedata:
        # regex match object can't be shelved so delete it
        del file[b"regexmatch"]
        myshelf["filedata"] += file
    myshelf.close()
    return None
