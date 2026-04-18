import os
import win32com.client as win32

path = os.getcwd()
videos = os.listdir()
shell = win32.Dispatch("Shell.Application")
no_title = 0

# returns metadata title from file
def game_title(file_path, title_id = 21):
    folder_path = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)
    folder = shell.NameSpace(folder_path)
    file = folder.ParseName(file_name)

    # sorts through metadata ids to get title id
    for i in range (300):
        metadata = folder.GetDetailsOf(None, i)
        if metadata == "Title":
            title_id = i
            continue
            
    title = folder.GetDetailsOf(file, title_id)

    # contigency if lack of title
    if title == "":
        title = "No Title"
    
    return title

# moves video to respective folder
def sort_video(title, file):
    # makes sure all characters in folder name are valid characters
    valid_title = ""
    for char in title:
        if char.isalnum() or char == " ":
            valid_title += char
    valid_title = valid_title.rstrip()
    valid_title += "\\"

    fso = win32.Dispatch("Scripting.FileSystemObject")
    dest_path = os.path.join(path, valid_title)
    og_path = os.path.join(path, file)

    # moves video to respective folder
    if os.path.isdir(dest_path):
        fso.MoveFile(og_path, dest_path)
    else:
        os.mkdir(dest_path)
        fso.MoveFile(og_path, dest_path)

    folder_path = shell.NameSpace(dest_path)
    print(f"{file} moved to the '{folder_path}' folder.\n")

    if str(folder_path) == "No Title":
        global no_title
        no_title += 1

# loops for each video file and calls both def
for file in videos:
    if file[-3:] == "mp4":
        video_path = os.path.join(path, file)
        title = game_title(video_path)
        sort_video(title, file)
        
print(f"-----------\n{no_title} file(s) moved to 'No Title' folder")
