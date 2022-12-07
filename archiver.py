#Program to archive projects
import pathlib
import os
import stat
from tkinter import *
from tkinter import scrolledtext
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime
from zipfile import ZipFile, ZIP_DEFLATED, ZipInfo
import shutil
from tkfilebrowser import askopendirnames
from shutil import Error
import re
from tkinterdnd2 import DND_FILES, TkinterDnD



def check_folder_struc(folder, folder_level):
    """ Takes file path to and checks folder structure for expected folders,
            returns loose files if found; else returns False for non-matching
            structure.
    """
    #list of folders to check for
    correct_struc = ['1. Photos','2. Communications','3. Scope & Quality',
                     '4. Time Mgt', '5. Cost & Procurement', '6. Risk Mgt',
                     '7. Drafting', '8. Design (Do not copy to client)']
    loose_files = []
    #check for folder level deeper than first two levels, stop searching here
    if folder_level > 1:
        return None
    #check if correct struc folders at current level
    folders_found = 0
    for k in folder.iterdir():
        if k.name in correct_struc:
            folders_found += 1
    #check for loose files
    for i in folder.iterdir():
        if not i.is_dir():
            loose_files.append(i)
        else:
            #check folder structure if at level below
            if folder_level == 1 and i.name not in correct_struc:
                if folders_found > 3:
                    print(F'folder {i} found not matching correct struc')
                    loose_files.append(i)
                else:
                    return False
            #check if recursive function call gives a wrong structure flag
            rec_check = check_folder_struc(i, folder_level + 1)
            if rec_check == False:
                print('passed folder not found to level up')
                return False
            elif rec_check:
                for j in rec_check:
                    loose_files.append(j)
    return loose_files
            

def delete_empty(folder):
    """Takes file path to project folder, recursively searches for empty folders
    and deletes. Returns list of deleted folders.
    """
    deleted_folder_list = []
    for i in folder.iterdir():
        #attempt to remove folder, raise OSerror if not empty
        if i.is_dir():
            delete_folder = i.name
            try:
                i.rmdir()
            except OSError:
                result = delete_empty(i)
                for j in result:
                    deleted_folder_list.append(j)
                #try again to remove folder after deleting subfolders
                try:
                    i.rmdir()
                except OSError:
                    delete_folder = None
            if delete_folder:
                deleted_folder_list.append(delete_folder)            
    return deleted_folder_list


def zip_single_files(folder, gui):
    """ Takes file path to folder and gui object and zips files meeting a
    certain criteria. Each file is zipped into its own archive.
    """
    skip_files = ['.pdf', '.jpg', '.jpeg', '.png', '.zip', '.zipx', '.msg']
    skip_folders = ['7.2. Inventor Files', '8.2. FEA', '8.3. STEELBeam',
                    '8.4. SpaceGASS', '2.1. Email In', '2.2. Email Out',
                    '1. Photos', '2.3. Document Transmittals']
    suffix_to_zip = ['.dwg', '.dwf', '.dwfx', 'dxf']
    zipped_num = 0
    failed_delete = []
    files_to_zip = []
    files = get_all_files(folder)
    for file in files:
        #zip files not in skipped folders meeting criteria
        parents_list = list(i.name for i in file.parents)
        if file.is_file() and not any(i in skip_folders for i in parents_list):
            # criteria to meet in order to zip file
            #(zip all non image files seperately)
            if (str.lower(file.suffix) in suffix_to_zip) or\
            (file.stat().st_size > 1E6 and str.lower(file.suffix)
            not in skip_files):
                zip_name = file.with_suffix('.zip')
                files_to_zip.append((file, zip_name))
            else:
                print(F"Skipped {file}")
    #initialize progress bar
    prog = progress_data(gui, len(files_to_zip))
    for file in files_to_zip:
        print(F"Zipping {file[0]}")
        with CallBackZipFile(file[1], 'w') as zfile:
            if file[0].stat().st_size > 1E+9:
                file_size = file[0].stat().st_size * 1E-9
                print(F"Compressing large file ({round(file_size, 2)} GB) -"
                      F"{file[0].name}")
                if prog:
                    prog.gui.updatetxt(F"Compressing large file"
                                       F"({round(file_size, 2)} GB) "
                                       F"- {file[0].name}")
                #initialize new sub-progress bar for zip progress
                prog.set_total_progress(file[0].stat().st_size)
                zfile.write(file[0], arcname = file[0].name,
                            compress_type = ZIP_DEFLATED,
                            compresslevel = 9, prog = prog)
            else:
                zfile.write(file[0], arcname = file[0].name,
                            compress_type = ZIP_DEFLATED, compresslevel = 9)
        #delete file after archiving
        result = del_file(file[0])
        if result is None:
            zipped_num += 1
            prog.increment_other()
        else:
            failed_delete.append(result)
    return zipped_num, failed_delete


def delete_temp(folder):
    """Takes file path to folder and deletes files with temp file suffixes"""
    temp_files = ['.bak', '.twl', '.dwl', '.dwl2', '.log', '.err']
    word_files = ['.doc', '.docx', '.xls', '.xlsx', '.xlsm']
    delete_num = 0
    failed_delete = []
    files = get_all_files(folder)
    for file in files:
        if file.is_file():
            if str.lower(file.suffix) in temp_files\
            or str.lower(file.name) == 'thumbs.db'\
            or (file.name[0:2] == '~$' and str.lower(file.suffix)
                in word_files):
                result = del_file(file)
                if result is None:
                    delete_num += 1
                else:
                    failed_delete.append(result)
    return delete_num, failed_delete



def zip_group_files(folder, prog, archive = False):
    """Takes a file path to folder and zips items in folder together.
    For folders in 'group_zip' sub folders will be ignored and zipped.
    Multiple files are zipped togther into an archive.
    """
    zip_list = []
    zipped_num = 0
    failed_delete = []
    #for folder in 'group zip' - this will zip all contents into one zip file regardless of subfolders
    group_zip = ['7.2. Inventor Files','8.2. FEA', '8.3. STEELBeam', '8.4. SpaceGASS', '8.5. Slabs & Footings' ]
    if any(i in folder.parts for i in group_zip) or archive:
        zip_name = folder.joinpath((folder.name + '.zip'))
        files = []
        for i in folder.rglob('*'):
            if i.is_file():
                files.append(i)
        print(zip_name)
        result = zip_together(files,zip_name,archive = True, prog = prog)
        zipped_num += result[0]
        failed_delete += result[1]
    #for folders not in 'group zip' subfolders will be maintained
    else:
        for file in folder.iterdir():
            if file.is_file():
                zip_list.append(file)
            elif file.is_dir():
                rec_zip = zip_group_files(file, prog)
                zipped_num += rec_zip[0]
                for i in rec_zip[1]:
                    failed_delete.append(i)
        zip_name = folder.joinpath((folder.name + '.zip'))
        if zip_list:
            result = zip_together(zip_list, zip_name, prog=prog)
            zipped_num += result[0]
            failed_delete += result[1]
    return zipped_num, failed_delete
                 

def copy_project(folder, gui):
    """Takes file path to folder and gui object and copies folder and contents
    to archive folder for archiving.
    """
    project = folder.name
    dest = convert_path("\\\\GRGSVRDATA\\Data\\Synergy\\Projects\\"
                        "Archived\\Unfiled Archived Projects").joinpath(project)
    #copytree fails for long file paths, catching this error
    #get number of files to set copy progress increment
    file_count = 0
    for path, dirs, files in os.walk(folder):
        file_count += len(files)        
    #initialize progress bar to store copy progress
    prog = progress_data(gui, file_count)
    try:
        copied_folder = shutil.copytree(folder,\
                                        dest, \
                                        symlinks = True,\
                                        copy_function = prog.increment_copy)
        return convert_path(copied_folder)
    except Error as err:
        #deal with files that created error
        for i in err.args:
            for j in i:
                new_copy_path = convert_path(j[0])
                new_dst_path = convert_path(j[1])
                if os.path.isfile(new_copy_path):
                    shutil.copy(new_copy_path, new_dst_path)
                else:
                    os.mkdir(new_dst_path)
                prog.increment_other()
    return convert_path(dest)

def get_group_folders(root_folder):
    """Takes file path to project folder and returns list of file paths
    to folders that the zip_group_files function is to be run on.
    """
    folder_list = ['1. Photos',['2. Communications', ['2.1. Email In', '2.2. Email Out']],['7. Drafting', ['7.2. Inventor Files']], ['8. Design (Do not copy to client)', ['8.2. FEA', '8.3. STEELBeam', '8.4. SpaceGASS']]]
    zip_folders = []
    for sub_folder in root_folder.iterdir():
        if sub_folder.is_dir():
            for i in folder_list:
                if type(i) == list:
                    for j in i[1]:
                        folder_path = sub_folder.joinpath(i[0]).joinpath(j)
                        if folder_path.is_dir():
                            zip_folders.append((folder_path, False))
                else:
                    folder_path = sub_folder.joinpath(i)
                    if folder_path.is_dir():
                        zip_folders.append((folder_path, False))
    #search for pack and go files not in group zip folders
    sub_folder = re.compile('\d\.?\d?\. ')
    for file in root_folder.rglob("*.ipj"):
        if '7.2. Inventor Files' not in file.parent.parts and\
           not sub_folder.search(file.parent.name):
            zip_folders.append((file.parent, True))
    return zip_folders

               
def move_loose(loose_files, copied_root):
    """Takes a list of file paths to loose files and moves to a folder
    called unsorted files
    """
    rel_loose_files = []
    #get file paths for loose files in copied folder
    for i in loose_files:
        filepath = copied_root
        for j in i.parts[i.parts.index(copied_root.name)+1:]:
                   filepath = filepath.joinpath(j)
        rel_loose_files.append(filepath)
    dest_path = copied_root.joinpath('Other Files')
    if not dest_path.exists():
        os.mkdir(dest_path)
    for file in rel_loose_files:
        try:
            shutil.move(file, dest_path.joinpath(file.name))
        except:
            print(F"Couldn't move {file.name}")
            continue

#checks Inventor Files folder for PDF drawing files
def check_for_drawings(folder):
    """Takes file path to 'Inventor Files' folder and searches for
    PDF drawing files meeting standard GRG naming format
    """
    moved_list = []
    #regular expression to match drawing number format
    dwg_format = re.compile('\d\d-\d\d\d\d-\d\d\d')
    files = get_all_files(folder)
    in_folders = ['7.2. Inventor Files']
    not_in_folders = ['7.3. Received Drawings', '7.1. Drawings']
    for f in files:
        match = dwg_format.match(f.name)
        if any(i in in_folders for i in f.parts) and str.lower(f.suffix) == '.pdf'\
           and match and (not any(i in not_in_folders for i in f.parts)):
            moved_list.append(f)
            #finding "Drawings Folder"
            for i in f.parents:
                if i.name == '7. Drafting':
                    dest = i.joinpath('7.1. Drawings')
            #move drawing to drawings folder
            print(f)
            try:        
                shutil.move(f, dest)
            except:
                print(F'Failed moving {f}')
    return moved_list
            
            
#zips .msg and image files not in correct
#folder together
        
def zip_loose_msg_img(folder):
    """Takes file path to project folder and zips all .msg and .img files.
    Only run after other zip functions have been run to catch any loose files.
    """
    image_list = ['.jpg', '.jpeg', '.png', '.gif']
    image_zip_list = {}
    msg_zip_list = {}
    failed_delete = []
    zipped_num = 0
    files = get_all_files(folder)
    for file in files:
        if str.lower(file.suffix) in image_list:
            if file.parent in image_zip_list.keys():
                image_zip_list[file.parent].append(file)
            else:
                image_zip_list[file.parent] = [file]
        elif str.lower(file.suffix) == '.msg':
            if file.parent in msg_zip_list.keys():
                msg_zip_list[file.parent].append(file)
            else:
                msg_zip_list[file.parent] = [file]
    #zip files together per folder
    if image_zip_list:
        for k in image_zip_list.keys():
            zip_dest = k.joinpath('Images.zip')
            result = zip_together(image_zip_list[k], zip_dest)
            zipped_num += result[0]
            failed_delete += result[1]
    if msg_zip_list:
        for k in msg_zip_list.keys():
            zip_dest = k.joinpath('MSG Files.zip')
            result = zip_together(msg_zip_list[k], zip_dest)
            zipped_num += result[0]
            failed_delete += result[1]
    return zipped_num, failed_delete 

    

def zip_together(files, zip_dest, archive = False, prog = None):
    """Takes a list of files and zips them together. Zip archive is named the
    same as the parent folder. If archive = True, creates a zip archive
    """
    failed_delete = []
    zipped_num = 0
    #check file sizes
    with CallBackZipFile(zip_dest, 'w') as zfile:
        for file in files:
            if file.stat().st_size > 1E+9 and prog:
                file_size = file.stat().st_size * 1E-9
                prog.gui.updatetxt(F"Compressing large file ({round(file_size, 2)} GB) - {file.name}")
                prog.set_total_progress(file.stat().st_size)
                zip_prog = prog
            else:
                zip_prog = None
            if archive:
                zfile.write(file, file.relative_to(zip_dest.parent),\
                            compress_type = ZIP_DEFLATED, compresslevel = 9,
                            prog = zip_prog)
            else:
                zfile.write(file, arcname = file.name,\
                            compress_type = ZIP_DEFLATED, compresslevel = 9,
                            prog = zip_prog)
            print(F"Zipped {file}")
            result = del_file(file)
            if result is None:
                zipped_num += 1
                if prog:
                    prog.increment_other()
            else:
                failed_delete.append(result)
    return zipped_num, failed_delete

#handles long paths and converts to pathlib path
def convert_path(path):
    if type(path) == pathlib.WindowsPath:
        path = path.as_posix()
    if '?' not in path[0:10]:
        new_path = os.fspath(pathlib.WindowsPath(path))
        if 'GRGSVRDATA' in new_path:
            new_path = u'\\\\?\\UNC\\' + new_path[2:]
        else:
            new_path = u'\\\\?\\' + new_path
        new_path = pathlib.WindowsPath(new_path)
    else:
        new_path = pathlib.WindowsPath(path)
    return new_path

def get_all_files(folder):
    file_list = []
    for path, dirs, files in os.walk(folder):
        for i in files:
            file = convert_path(F"{path}\\{i}")
            file_list.append(file)
    return file_list

#attempts deletion, changes to read only if fails once
def del_file(file):
    try:
        file.unlink()
        print(F"Deleted {file}")
        return None
    except PermissionError:
        print(F"Couldn't delete {file}")
        #change file to read only
        os.chmod(file, stat.S_IWRITE)
        #attempt deletion again
        try:
            file.unlink()
            print(F"Deleted {file}")
        except PermissionError:
            return file


class CallBackZipFile(ZipFile):
    """Modified zipfile class to add a callback function for zipfile writing
        to allow progress bar to update for large zip files"""
    
    def write(self, filename, arcname=None,
              compress_type=None, compresslevel=None, prog = None):
        """Put the bytes from filename into the archive under the name
        arcname."""
        if not self.fp:
            raise ValueError(
                "Attempt to write to ZIP archive that was already closed")
        if self._writing:
            raise ValueError(
                "Can't write to ZIP archive while an open writing handle exists"
            )

        zinfo = ZipInfo.from_file(filename, arcname,
                                  strict_timestamps=self._strict_timestamps)

        if zinfo.is_dir():
            zinfo.compress_size = 0
            zinfo.CRC = 0
        else:
            if compress_type is not None:
                zinfo.compress_type = compress_type
            else:
                zinfo.compress_type = self.compression

            if compresslevel is not None:
                zinfo._compresslevel = compresslevel
            else:
                zinfo._compresslevel = self.compresslevel

        if zinfo.is_dir():
            with self._lock:
                if self._seekable:
                    self.fp.seek(self.start_dir)
                zinfo.header_offset = self.fp.tell()  # Start of header bytes
                if zinfo.compress_type == ZIP_LZMA:
                # Compressed data includes an end-of-stream (EOS) marker
                    zinfo.flag_bits |= 0x02

                self._writecheck(zinfo)
                self._didModify = True

                self.filelist.append(zinfo)
                self.NameToInfo[zinfo.filename] = zinfo

                self.fp.write(zinfo.FileHeader(False))
                self.start_dir = self.fp.tell()
        else:
            with open(filename, "rb") as src, self.open(zinfo, 'w') as dest:
                copyfileobj(src, dest, 1024*8, callback = prog)



def copyfileobj(fsrc, fdst, length=0, callback = None):
    """custom copyfileobj with callback for progress"""
    # Localize variable access to minimize overhead.
    if not length:
        length = shutil.COPY_BUFSIZE
    fsrc_read = fsrc.read
    fdst_write = fdst.write
    prog = 0
    while True:
        buf = fsrc_read(length)
        if not buf:
            break
        fdst_write(buf)
        prog += len(buf)
        if callback:
            callback.increment_zip(prog)


        

class progress_data:

    def __init__(self, gui, num_files = None):
        self.progress = 0
        #window to update progress with
        self.gui = gui
        if num_files:
            self.increment = 100/num_files
        else:
            self.increment = 100

    #method to increment progress bar for copying    
    def increment_copy(self, src, dst):
        self.progress += self.increment
        self.gui.updateprogress(self.progress)
        #catch file path too long errors (probably not required)
        '''
        if len(src) > 255 or len(dst) > 255:
            new_copy_path = os.fspath(pathlib.WindowsPath(src))
            new_dst_path = os.fspath(pathlib.WindowsPath(dst))
            if len(src) > 255:
                new_copy_path = u'\\\\?\\UNC\\' + new_copy_path[2:]
                print('source path name too long')
                print(new_copy_path)
            if len(dst) > 255:
                new_dst_path = u'\\\\?\\' + new_dst_path
                print('dest path name too long')
                print(new_dst_path)
            shutil.copy2(new_copy_path, new_dst_path)
        else:'''
        shutil.copy2(src, dst)
    
        
    #method to increment progress bar for other functions
    def increment_other(self):
        self.progress +=self.increment
        self.gui.updateprogress(self.progress)

    def set_complete(self):
        self.progress = 100
        self.gui.updateprogress(self.progress)

    def increment_zip(self, sofar):
        self.zip_progress = sofar/self.total * 100
        self.gui.updatesubprogress(self.zip_progress)

    def set_total_progress(self, total):
        self.total = total




#GUI


class igui:

    def __init__(self, master):
        self.master = master            #master = root = tk()
        self.master.title("Project Folder Archiver")   #master.title = root.title = tk().title
        
        #initializing tkk frame to hold widgets, sets size to expand and padding
        self.mainframe = ttk.Frame(master, padding = '20 12 2 2')
        self.mainframe.grid(column = 0, row = 0, sticky = (N,S,E,W))
        self.mainframe.columnconfigure(0, weight = 1)
        self.mainframe.rowconfigure(0, weight = 1)
        


        #labels and buttons
        self.NewButton = ttk.Button(self.mainframe, text = 'Add Project Folder', width=40, command = self.select_folder)
        self.NewButton.grid(column = 1, row = 1, sticky = W)
        self.RemoveButton = ttk.Button(self.mainframe, text = 'Remove Project Folder', width=40, command = self.remove_list)
        self.RemoveButton.grid(column = 2, row = 1, sticky = W)
        self.HelpButton = ttk.Button(self.mainframe, width=10, text='?', command = self.help)
        self.HelpButton.grid(column = 3, row = 1, sticky = E)
        self.GoButton = ttk.Button(self.mainframe, text = 'Archive!', width = 40, command = self.archive)
        self.GoButton.grid(column = 1, row = 3, sticky = W)
        #listbox
        self.ListBox = Listbox(self.mainframe, width = 100)
        self.ListBox.grid(column = 1, row = 2, sticky = N, columnspan = 3)
        self.dnd_message = 'Drag and drop project folders here...'
        self.ListBox.insert(END, self.dnd_message)

        #Listbox drag and drop
        self.ListBox.drop_target_register(DND_FILES)
        self.ListBox.dnd_bind('<<Drop>>', self.lbox_dnd)

        
        for child in self.mainframe.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def lbox_dnd(self, e):
        print(e.data)
        path_list = []    
        for i in re.split('[{}]', e.data):
            if os.path.exists(i):
                if os.path.isdir(i):
                    path_list.append(i)
                else:
                    continue
            else:
                resplit = i.split()
                for k in resplit:
                    if os.path.isdir(k):
                        path_list.append(k)
        if path_list:
            if self.ListBox.size() == 1 and self.ListBox.get(0) == self.dnd_message:
                self.ListBox.delete(0, END)
            for j in path_list:
                self.ListBox.insert(END, j)

            

    def select_folder(self):
        filetypes = (('Message files', '*.msg'),('All files', '*.*'))
        filenames = askopendirnames(title = 'Select project folders', initialdir = 'P:\\')
        if filenames:
            print(self.ListBox.size())
            print(self.ListBox.get(1))
            if self.ListBox.size() == 1 and self.ListBox.get(0) == self.dnd_message:
                self.ListBox.delete(0, END)
            for num, item in enumerate(filenames, start = 1):
                self.ListBox.insert(num, item)

    def archive(self):
        folder_list = list(self.ListBox.get(0, self.ListBox.size()))
        #convert to pathlib
        for ind, item in enumerate(folder_list):
            folder_list[ind] = convert_path(item)
        #create new window
        self.mwindow = Toplevel(self.master)
        #make loading window transiet ie. drawn over top of main window
        self.mwindow.transient()
        #pass window to loadwindow class
        self.app = loadwindow(self.mwindow)
        #change focus to loading window
        self.mwindow.geometry("+%d+%d" % (self.master.winfo_x(), self.master.winfo_y()))
        self.mwindow.grab_set()
        #now run functions

        #log of number of files modified
        deleted_log = 0
        zipped_log = 0
        loose_log = 0
        moved_log = 0
        failed_delete = []
        for folder in folder_list:
            print('run loose files')
            loose_files = check_folder_struc(folder, 0)
            print(loose_files)
            #check if folder structure is correct
            if loose_files == False:
                print('skip')
                self.app.updatetxt(F"Skipped {folder.name} due to incorrect folder structure...")
                continue
            else:
                self.app.updatetxt(F"Copying {folder.name} to archived folder for archiving...")
                try:
                    copied_folder = copy_project(folder, self.app)
                except FileExistsError:
                    self.app.updatetxt(F"Project folder {folder} already exists in archived folder.")
                    continue
                self.app.updatetxt("Done!")
                if loose_files:
                    self.app.updatetxt("Moving loose files...")
                    move_loose(loose_files, copied_folder)
                    loose_log += len(loose_files)
                    self.app.updatetxt("Done!")
                self.app.updatetxt("Checking for drawing PDFs...")
                moved_log += len(check_for_drawings(copied_folder))
                self.app.updatetxt("Done!")                 
                self.app.updatetxt("Clearing empty folders...")
                deleted_folders = delete_empty(copied_folder)
                deleted_log += len(deleted_folders)
                self.app.updatetxt("Done!")
                self.app.updatetxt("Deleting temp files...")
                t = delete_temp(copied_folder)
                deleted_log += t[0]
                failed_delete += t[1]
                self.app.updatetxt('Done!')
                self.app.updatetxt("Zipping single files...")
                t = zip_single_files(copied_folder, self.app)
                zipped_log += t[0]
                failed_delete += t[1]
                self.app.updatetxt('Done!')
                self.app.updatetxt("Zipping group files...")
                fold = get_group_folders(copied_folder)
                #initialize progress bar
                to_zip = sum(len(list(i[0].rglob('*')))for i in fold)
                prog = progress_data(self.app, to_zip)
                for i in fold:
                    t = zip_group_files(i[0], prog, i[1])
                    zipped_log += t[0]
                    failed_delete += t[1]
                #delete empty folders again after group zipping
                delete_empty(copied_folder)
                self.app.updatetxt('Done!')
                self.app.updatetxt("Zipping loose image and msg files...")
                prog.set_complete()
                r = zip_loose_msg_img(copied_folder)
                zipped_log += r[0]
                failed_delete += r[1]
                #clear thumbs.db files if created
                delete_temp(copied_folder)
                self.app.updatetxt('Archived!')
        self.app.updatetxt(F"Deleted {deleted_log} empty folders and temp files.")
        self.app.updatetxt(F"Zipped {zipped_log} files.")
        self.app.updatetxt(F"Moved {loose_log} files.")
        self.app.updatetxt(F"Moved {moved_log} drawing PDF files.")
        for i in failed_delete:
            self.app.updatetxt(F"Failed deleting {i}")
        self.app.ok_button()
        self.app.lock_state()
        self.ListBox.delete(0, END)
        self.ListBox.insert(END, self.dnd_message)
        self.mwindow.update()
        os.startfile("\\\\GRGSVRDATA\\Data\\Synergy\\Projects\\Archived\\Unfiled Archived Projects")    

    def help(self):
        self.helpmessage = messagebox.showinfo(message = """This app archives multiple selected project folders by performing the following functions:
                                               \n - Checks project folder structure matches GRG standard
                                               \n - Copies folder to P:\\Archived\\Unfiled Archived Projects
                                               \n - Copies loose files not in standard folder structure to a folder called Other Files
                                               \n - Zips files >1mb in size but not PDFs and zips inventor files, photos and communcations together
                                               \n - Deletes empty folders""")

    def remove_list(self):
        for i in self.ListBox.curselection():
            if i != 'Drag and drop project folders here...':
                self.ListBox.delete(i)
        if self.ListBox.size() == 0:
            self.ListBox.insert(END, self.dnd_message)
        

        
    

#loading window
class loadwindow:
    def __init__(self, master):
        self.master = master
        self.log = scrolledtext.ScrolledText(self.master, height = 20, width = 75, wrap = WORD)
        self.log.grid_configure(padx = 20, pady = 20)
        self.log.grid(row = 4, column = 1, sticky = N)
        self.log.insert(END, "Archiving...")

        #progress bar
        self.progress = ttk.Progressbar(self.master, orient = 'horizontal', \
                                        mode = 'determinate', length = 380)
        self.progress.grid_configure(padx = 30, pady = (2, 10), ipady = 10)
        self.progress.grid(column = 1, row = 2, sticky = N)

        #sub progress bar
        self.subprogress = ttk.Progressbar(self.master, orient = 'horizontal', \
                                        mode = 'determinate', length = 380)
        self.subprogress.grid_configure(padx = 30, pady = (1, 1), ipady = 1)
        self.subprogress.grid(column = 1, row = 3, sticky = N)
        self.subprogress.grid_remove()
        

        #progress bar label
        self.prog_num = StringVar()
        self.prog_num.set("0%")
        self.plabel = ttk.Label(self.master, textvariable = self.prog_num, font = (None, 12))
        self.plabel.grid_configure(padx = 5, pady = (30, 5))
        self.plabel.grid(column = 1, row = 1, sticky = N)

    def updateprogress(self, percent):
        self.progress['value'] = percent
        self.prog_num.set(F"{round(percent)}%")
        self.subprogress.grid_remove()
        self.master.update()


    def updatetxt(self, text):
        newtext = '\n' + text
        self.log.insert(END, newtext)
        self.log.see("end")
        self.master.update()

    def ok_button(self):
        self.NewButton = ttk.Button(self.master, text = 'OK', width=40, command = self.master.destroy)
        self.NewButton.grid(row = 5, column = 1, sticky = N)
        self.NewButton.grid_configure(padx = 20, pady = 20)

    def lock_state(self):
        self.log.configure(state = 'disabled')

    def updatesubprogress(self, percent):
        self.subprogress['value'] = percent
        self.subprogress.grid()
        self.master.update()
        



root=TkinterDnD.Tk()
app=igui(root)
root.mainloop()

