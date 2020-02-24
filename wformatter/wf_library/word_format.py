from os import path, remove, rename
from shutil import copyfile, make_archive, rmtree
from zipfile import ZipFile
from datetime import datetime
from abc import abstractmethod, ABC

class WordFormatter(object):
    
    def __init__(self, base_file: str, data: dict, logger: Logger):
        '''
        This class will create a copy of the word file and replace the text content of the word file defined in the 
        data keys with the values.

        Initialize the Word Formatter instance with the base file that takes full path of the docx file, 
        data of key value pairs, and logger instance that has log method to log the error events.

        base_file - string -> Full path of the docx file.
        data - dictionary -> Contains key value pair data. Keys will be replaced by values.
        logger - Instance of Logger's subclasses -> Used to log the error messages that occur while reading/writing xml file.
        '''

        # Get file url and split it to get folder url and file names
        splitted_file = path.split(base_file)
        # Save data, settings
        self.__base_file = base_file
        self.__data = data
        self.__file_path = splitted_file[0]
        self.__file = splitted_file[1]
        self.__file_copy = "copy_" + self.__file
        self.__file_copy_path = self.__file_path + "\\" + self.__file_copy
        self.__file_copy_zip = path.splitext(self.__file_copy)[0] + ".zip"
        self.__file_copy_zip_path = self.__file_path + "\\" + self.__file_copy_zip
        self.__extract_folder_name = splitted_file[0] + r"\extract_" + path.splitext(self.__file)[0]
        self.__xml_file = self.__extract_folder_name + "\\word\\" + "document.xml"
        self.__logger = logger

    
    def __make_file_copy(self):
        # Make a file copy to work on. This copy will be given to the user.
        copyfile(self.__base_file, self.__file_copy_path)

    
    def __rename_file_copy(self, zip_to_docx=False):
        if zip_to_docx:
            # Rename file copy extension to docx.
            rename(self.__file_copy_zip_path, self.__file_copy_path)
        else:    
            # Rename file copy extension to zip so that it can be extracted.
            rename(self.__file_copy_path, self.__file_copy_zip_path)


    def __extract_file_copy(self):
        # Extract file copy zip to a folder.
        with ZipFile(self.__file_copy_zip_path, mode='r') as zip:
            zip.extractall(self.__extract_folder_name)
    

    def __replace_file_content(self):
        # Read the contents of the xml file.
        content = self.__read_xml_file_content()
        
        # Replace the content of the xml file. 
        # Placeholders are replaced with the data.
        for key, value in self.__data.items():
            content = content.replace(key, value)
        
        # Write the new content of the file.
        self.__write_xml_file_content(content)
                

    def __read_xml_file_content(self):
        content = ""
        time = datetime.now().strftime("%m/%d/%Y, %H:%M:%S")

        try:
            file = open(self.__xml_file, 'r')
            content = file.read()
            file.close()
        except FileNotFoundError as not_found_error:
            self.__logger.log("Error reading file: " + not_found_error)
            self.__logger.log("Time: " + time)
        except FileExistsError as exists_error:
            self.__logger.log("Error reading file: " + exists_error)
            self.__logger.log("Time: " + time)
        except:
            self.__logger.log("Error reading file. FileError at " + time)

        return content


    def __write_xml_file_content(self, content):
        time = datetime.now().strftime("%m/%d/%Y, %H:%M:%S")

        try:
            file = open(self.__xml_file, 'w')
            file.write(content)
            file.close()
        except FileNotFoundError as not_found_error:
            self.__logger.log("Error writing file: " + not_found_error)
            self.__logger.log("Time: " + time)
        except FileExistsError as exists_error:
            self.__logger.log("Error writing file: " + exists_error)
            self.__logger.log("Time: " + time)
        except:
             self.__logger.log("Error writing file. FileError at " + time)


    def __delete_old_zip_file(self):
        # Remove old zip file.
        remove(self.__file_copy_zip_path)


    def __compress_edited_file(self):
        # Compress the folder into a zip file so that it can be renamed.
        zip_file_name = path.splitext(self.__file_copy_zip_path)
        make_archive(zip_file_name[0], "zip", self.__extract_folder_name)
    

    def __delete_extract_tree(self):
        # Delete the extracted folder.
        rmtree(self.__extract_folder_name)


    def execute_formatting(self):
        '''
        Starts the execution of formatting the document.
        '''
        self.__make_file_copy()
        self.__rename_file_copy()
        self.__extract_file_copy()
        self.__replace_file_content()
        self.__delete_old_zip_file()
        self.__compress_edited_file()
        self.__delete_extract_tree()
        self.__rename_file_copy(True)



class Logger(ABC):

    @abstractmethod
    def log(self, message):
        pass

