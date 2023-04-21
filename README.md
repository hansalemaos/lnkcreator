# Creates .lnk files on Windows (with Admin rights if desired)

## pip install lnkcreator



```python
	
		
from lnkcreator import create_shortcut
create_shortcut(
    shortcut_path=r"C:\Users\hansc\Downloads\PJeOffice.lnk",
    target=r"C:\Users\hansc\Downloads\PJeOffice.exe",
    arguments=(),
    minimized_maximized_normal='normal',  #
    asadmin=True,  # enables the admin check box
    hotkey='ctrl+x',
    working_dir=None, # dir of target will be used 
)



Creates a Windows shortcut (.lnk) file at the specified path with the specified properties.

Args:
	shortcut_path (str): The path where the shortcut file will be created.
	target (str): The path of the target file or application that the shortcut will point to.
	arguments (list): A list of arguments to be passed to the target file or application.
	hotkey (str, optional): The hotkey combination to activate the shortcut. Defaults to "".
	working_dir (Union[str, None], optional): The working directory for the target file or application. Defaults to None.
	minimized_maximized_normal (str, optional): The window state of the target application when the shortcut is activated.
		Possible values are "minimized", "maximized", or "normal". Defaults to "minimized".
	asadmin (bool, optional): If True, the shortcut will be created with administrative privileges. Defaults to False.

Returns:
	str: The JavaScript content used to create the shortcut.

Raises:
	OSError: If the shortcut file cannot be created.



```