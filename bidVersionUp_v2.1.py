import openpyxl, os, sys, shutil, warnings


def main():
    #get file path from user
    filepath = input("Filepath: ")

    #parse file, check file extension, and extract version number
    file = file_parse(filepath)
    try:
        if str(file[2]) in [".xls", ".xlsm"]:
            bid = file[1].split("_")
            print("Bid: " , "_".join(bid))
        else:
            sys.exit()
    except:
        #print(file[2])
        sys.exit("Not an Excel File")

    else:
        #open bid in excel, update bid version + save versioned up file.
        bid_file = version_up(bid)
        #bid_file[0] - Bid File name without extension
        #bid_file[1] - Bid Version
        try:
            warnings.simplefilter(action='ignore', category=UserWarning) #ignore warning about Data Validation
            new_bid = openpyxl.load_workbook(filename=filepath, read_only=False, keep_vba=True) #updated line in v2.1 to include keep_vba
        except FileNotFoundError:
            print("File Not Found, check path and try again")
            sys.exit()
        else:

            breakdown = new_bid['Brkdwn']

            breakdown.cell(row = 4, column = 10).value = "Bid_v" + str(bid_file[1])
            print("Updating Bid Version to: ", breakdown["J4"].value)
            print("Next Bid Version: Bid_v", bid_file[1] )
            print("file_root_path: ",file[3])
            new_bid.save(file[3]+"/"+bid_file[0]+file[2])
            print("New Bid Located: ", file[3]+"/"+bid_file[0]+file[2])

        #move previous version  to _Old directory
        new_path = new_Directory(filepath)
        print("New Path: ",new_path)
        move = shutil.move(filepath,new_path)
        print("File Moved: ",move)


    finally:
        pass

def file_parse(file):

    try:
        # this will return a tuple of root and extension
        split_tup = os.path.splitext(file)
        # extract the file name and extension
        file_path = split_tup[0]
        file_extension = split_tup[1]
        file_name = file_path.split("/")[-1]
        file_root_path = "/".join(file_path.split("/")[0:len(file_path.split("/"))-1])

        return [file_path,file_name,file_extension,file_root_path]

    except:
        print("Error checking file extension")

def new_Directory(filepath):
    folders = filepath.split("/")
    folders[len(folders)-2] = "_OLD"
    newDirectory = "/".join(folders)
    return newDirectory

def version_up(bid):
    version = 0
    version_position = 0
    version_up_bid = ""
    for i in range(len(bid)):
        if bid[i][0].lower() == "v":
            version = int(bid[i][1:4])
            x = i
    version +=1
    version = str(version).zfill(3) #zfill to add leading zeros
    bid[x] = version #get version number from file name using bid position from loop
    version_up_bid = "_".join(bid)
    return [version_up_bid,version]



main()
