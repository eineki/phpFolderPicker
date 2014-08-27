<?php
/**
 *
 * FolderPicker
 * Version 0.1.0
 *
 * FolderPicker is a php module leveraging on Shell.Application ActiveX object for choosing a
 * file or a directory using a php command line script.
 *
 *   The file should be a
 *
 * Usage Example:
 *
 * $selection = new FolderPicker();
 * $selection->setRootFolder(dirPicker::ssfDRIVES)
 *           ->setTitle("Installation Directory")
 *           ->setFlags(dirPicker::BIF_RETURNONLYFSDIRS);
 * $base = $selection->choose();
 * echo "The chosen path is {$base->path}\n"
 *
 * References:
 * msdn.microsoft.com Shell reference                         http://msdn.microsoft.com/en-us/library/windows/desktop/ff521731(v=vs.85).aspx
 * msdn.microsoft.com ShellSpecialFolderConstants enumeration http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx
 * msdn.microsoft.com Shell.BrowseForFolder method            http://msdn.microsoft.com/en-us/library/windows/desktop/bb774065%28v=vs.85%29.aspx
 * msdn.microsoft.com BROWSEINFO structure                    http://msdn.microsoft.com/en-us/library/windows/desktop/bb773205%28v=vs.85%29.aspx
 * msdn.microsoft.com Folder object                           http://msdn.microsoft.com/en-us/library/windows/desktop/bb787868%28v=vs.85%29.aspx
 * msdn.microsoft.com FolderItem object                       http://msdn.microsoft.com/en-us/library/windows/desktop/bb787810%28v=vs.85%29.aspx
 */

namespace liberosoftware_net\windows\cli\com;

class FolderPicker_Browse_Folder_Exception extends \Exception {}
class FolderItem_Unknown_Property_Exception extends \Exception {}

class FolderPicker {
	private $shell;
	private $title = "Choose a directory";
	private $root = 0;
	private $flag = 0;

	/**
	 * Windows desktop—the virtual folder that is the root of the namespace.
	 */
	const ssfDESKTOP = 0x0;

	/**
	 * File system directory that contains the user's program groups (which are
	 * also file system directories).
	 *
	 * A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs.
	 */
	const ssfPROGRAMS = 0x2;

	/**
	 * Virtual folder that contains icons for the Control Panel applications.
	 */
	const ssfCONTROLS = 0x03;

	/**
	 * Virtual folder that contains installed printers.
	 */
	const ssfPRINTERS = 0x4;

	/**
	 * File system directory that serves as a common repository for a user's documents.
	 *
	 * A typical path is C:\Users\username\Documents.
	 */
	const ssfPERSONAL = 0x5;

	/**
	 * File system directory that serves as a common repository for the user's favorite URLs.
	 *
	 * A typical path is C:\Documents and Settings\username\Favorites.
	 */
	const ssfFAVORITES = 0x6;

	/**
	 * File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user first logs into their profile after a reboot.
	 *
	 * A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\StartUp.
	 */
	const ssfSTARTUP = 0x7;

	/**
	 * File system directory that contains the user's most recently used documents.
	 *
	 * A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Recent.
	 */
	const ssfRECENT = 0x8;
	 
	/**
	 * File system directory that contains Send To menu items.
	 *
	 * A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\SendTo.
	 */
	const ssfSENDTO = 0x9;

	/**
	 * Virtual folder that contains the objects in the user's Recycle Bin.
	 */
	const ssfBITBUCKET = 0xa;

	/**
	 * File system directory that contains Start menu items.
	 *
	 * A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Start Menu.
	 */
	const ssfSTARTMENU = 0xb;

	/**
	 * File system directory used to physically store the file objects that are displayed on the desktop.
	 * It is not to be confused with the desktop folder itself, which is a virtual folder.
	 *
	 * A typical path is C:\Documents and Settings\username\Desktop.
	 */
	const ssfDESKTOPDIRECTORY = 0x10;
	 
	/**
	 * My Computer — the virtual folder that contains everything on the local computer: <em>storage devices, printers, and Control Panel.</em>
	 *
	 * This folder can also contain mapped network drives.
	 */
	const ssfDRIVES = 0x11;

	/**
	 * Network Neighborhood—the virtual folder that represents the root of the
	 * network namespace hierarchy.
	 */
	const ssfNETWORK = 0x12;

	/**
	 * A file system folder that contains any link objects in the My Network Places virtual folder.
	 * It is not the same as ssfNETWORK, which represents the network namespace root.
	 *
	 * A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Network Shortcuts.
	 */
	const ssfNETHOOD = 0x13;

	/**
	 * Virtual folder that contains installed fonts.
	 *
	 * A typical path is C:\Windows\Fonts.
	 */
	const ssfFONTS = 0x14;

	/**
	 * File system directory that serves as a common repository for document templates.
	 */
	const ssfTEMPLATES = 0x15;

	/**
	 * File system directory that contains the programs and folders that appear on the Start menu for all users.
	 *
	 * Valid only for Windows NT systems.
	 *
	 * A typical path is C:\Documents and Settings\All Users\Start Menu.
	 */
	const ssfCOMMONSTARTMENU = 0x16;

	/**
	 * File system directory that contains the directories for the common program groups that appear on the Start menu for all users.
	 *
	 * Valid only for Windows NT systems.
	 *
	 * A typical path is C:\Documents and Settings\All Users\Start Menu\Programs.
	 */
	const ssfCOMMONPROGRAMS = 0x17;

	/**
	 * File system directory that contains the programs that appear in the Start up folder for all users.
	 *
	 * Valid only for Windows NT systems.
	 *
	 * A typical path is C:\Documents and Settings\All Users\Microsoft\Windows\Start Menu\Programs\StartUp.
	 */
	const ssfCOMMONSTARTUP = 0x18;

	/**
	 * File system directory that contains files and folders that appear on the
	 * desktop for all users.
	 *
	 * Valid only for Windows NT systems.
	 *
	 * A typical path is C:\Documents and Settings\All Users\Desktop.
	 */
	const ssfCOMMONDESKTOPDIR = 0x19;

	/**
	 * File system directory that serves as a common repository for application-specific data.
	 *
	 * Version 4.71.
	 *
	 * A typical path is C:\Documents and Settings\username\Application Data.
	 */
	const ssfAPPDATA = 0x1a;

	/**
	 * File system directory that contains any link objects in the Printers virtual folder.
	 *
	 * A typical path is C:\Users\username\AppData\Roaming\Microsoft\Windows\Printer Shortcuts.
	 */
	const ssfPRINTHOOD = 0x1b;

	/**
	 * File system directory that serves as a data repository for local (non-roaming) applications.
	 *
	 * Version 5.0.
	 *
	 * A typical path is C:\Users\username\AppData\Local.
	 */
	const ssfLOCALAPPDATA = 0x1c;

	/**
	 *  File system directory that corresponds to the user's non-localized Startup program group.
	 */
	const ssfALTSTARTUP = 0x1d;

	/**
	 * File system directory that corresponds to the non-localized Startup program group for all users.
	 *
	 * Valid only for Windows NT systems.
	 */
	const ssfCOMMONALTSTARTUP = 0x1e;

	/**
	 * File system directory that serves as a common repository for the favorite URLs shared by all users.
	 *
	 *  Valid only for Windows NT systems.
	 */
	const ssfCOMMONFAVORITES = 0x1f;

	/**
	 * File system directory that serves as a common repository for temporary Internet files.
	 *
	 * A typical path is C:\Users\username\AppData\Local\Microsoft\Windows\Temporary Internet Files.
	 */
	const ssfINTERNETCACHE = 0x20;

	/**
	 * File system directory that serves as a common repository for Internet cookies.
	 *
	 * A typical path is C:\Documents and Settings\username\Application Data\Microsoft\Windows\Cookies.
	 */
	const ssfCOOKIES = 0x21;

	/**
	 * File system directory that serves as a common repository for Internet history items.
	 */
	const ssfHISTORY = 0x22;

	/**
	 * Application data for all users.
	 *
	 * Version 5.0.
	 *
	 * A typical path is C:\Documents and Settings\All Users\Application Data.
	 */
	const ssfCOMMONAPPDATA = 0x23;

	/**
	 * Windows directory. This corresponds to the %windir% or %SystemRoot% environment variables.
	 *
	 * Version 5.0.
	 *
	 * A typical path is C:\Windows.
	 */
	const ssfWINDOWS = 0x24;

	/**
	 * The System folder.
	 *
	 * Version 5.0.
	 *
	 * A typical path is C:\Windows\System32.
	 */
	const ssfSYSTEM = 0x25;

	/**
	 * Program Files folder.
	 *
	 * Version 5.0.
	 *
	 * A typical path is C:\Program Files.
	 */
	const ssfPROGRAMFILES = 0x26;

	/**
	 * My Pictures folder.
	 *
	 * A typical path is C:\Users\username\Pictures.
	 *
	 */
	const ssfMYPICTURES = 0x27;
	 
	/**
	 * User's profile folder.
	 *
	 * Version 5.0.
	 */
	const ssfPROFILE = 0x28;
	 
	/**
	 * System folder
	 *
	 * Version 5.0.
	 *
	 * A typical path on a 32-bit system is C:\Windows\System32.
	 * A typical path on a 64-bit envinronment is C:\Windows\Syswow32.
	 */
	const ssfSYSTEMx86 = 0x29;

	/**
	 * Program Files folder.
	 *
	 * Version 6.0.
	 *
	 * A typical path is C:\Program Files,
	 * on a 64-bit system: C:\Program Files (X86).
	 */
	const ssfPROGRAMFILESx86   = 0x30;

	// Flags specify the options for the dialog box. This member can be 0
	// or a combination of the following constant

	/**
	 * Only return file system directories. If the user selects forders that are not part of the file system, the OK button is grayed.
	 */
	const BIF_RETURNONLYFSDIRS = 0x1;

	/**
	 * Do not include network folders below the domain level in the dialog box's tree view control.
	 */
	const BIF_DONTGOBELOWDOMAIN = 0x2;

	/**
	 * Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box.
	 *
	 * This flag is not supported when BIF_NEWDIALOGSTYLE is specified.
	 *
	 * UNSUPPORTED
	 */
	const BIF_STATUSTEXT = 0x4;

	/**
	 * Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the OK button is grayed.
	 */
	const BIF_RETURNFSANCESTORS = 0x8;

	/**
	 * Include an edit control in the browse dialog box that allows the user to type the name of an item.
	 *
	 * UNSUPPORTED
	 * Version 4.71.
	 */
	const BIF_EDITBOX = 0x10;

	/**
	 * If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message.
	 *
	 * This flag is ignored if BIF_EDITBOX is not specified.
	 *
	 * Version 4.71.
	 */
	const BIF_VALIDATE = 0x20;

	/**
	 * Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities, including: drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands.
	 *
	 * Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_NEWDIALOGSTYLE is passed.
	 *
	 * Version 5.0.
	 */
	const BIF_NEWDIALOGSTYLE = 0x40;

	/**
	 * Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE.
	 *
	 * Version 5.0.
	 */
	const BIF_USENEWUI = 0x50;

	/**
	 * The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set. If any of these three flags are not set, the browser dialog box rejects URLs.
	 * Even when these flags are set, the browse dialog box displays URLs only if the folder that contains the selected item supports URLs.
	 * When the folder's IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not display the URL.
	 *
	 * Version 5.0.
	 */
	const BIF_BROWSEINCLUDEURLS = 0x80;

	/**
	 * When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box. BIF_EDITBOX overrides this flag.
	 *
	 * Version 6.0.
	 */
	const BIF_UAHINT = 0x100;

	/**
	 * Do not include the New Folder button in the browse dialog box.
	 *
	 * Version 6.0.
	 */
	const BIF_NONEWFOLDERBUTTON = 0x200;

	/**
	 * When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
	 *
	 * Version 6.0.
	 */
	const BIF_NOTRANSLATETARGETS = 0x400;

	/**
	 * Only return computers. If the user selects anything other than a computer, the OK button is grayed.
	 */
	const BIF_BROWSEFORCOMPUTER = 0x1000;

	/**
	 * Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed.
	 *
	 * In Windows XP and later systems, the best practice is to use a Windows XP-style dialog, setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS).
	 */
	const BIF_BROWSEFORPRINTER = 0x2000;

	/**
	 * The browse dialog box displays files as well as folders.
	 *
	 * Version 4.71.
	 */
	const BIF_BROWSEINCLUDEFILES = 0x4000;

	/**
	 * The browse dialog box can display sharable resources on remote systems. This is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set.
	 *
	 * Version 5.0.
	 */
	const BIF_SHAREABLE = 0x8000;

	/**
	 * Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed.
	 *
	 * Windows 7 and later.
	 */
	const BIF_BROWSEFILEJUNCTIONS = 0x10000;

	/**
	 * Inject a shell object (useful for testing purpose)
	 *
	 * @param \COM $shell
	 *
	 * @return FolderPicker
	 */
	public function setShell($shell) {
		$this->_shell = $shell;
		return $this;
	}

	/**
	 * Private function used internally to retrieve the instantiated shell.
	 * It create a default one if no shell has previously given.
	 *
	 * @return \COM
	 */
	private function _getShell() {
		if (!isset($this->_shell)) {
			$this->setShell(new \COM('Shell.Application'));
		}
		return $this->_shell;
	}

	/**
	 * This function show the file selection dialog and
	 *
	 * @throws FolderPicker_Browse_Folder_Exception
	 * @return FolderItem
	 */
	public function choose() {
		$shell = $this->_getShell();
		try {
			$selection = $shell->BrowseForFolder(0, $this->title, $this->flag, $this->root);
		} catch (COM_Exception $e) {
			throw new FolderPicker_Browse_Folder_Exception(
					$e->getMessage(),
					$e->getCode()
			);
		}
		return new FolderItem($selection);
	}

	public function setTitle($string) {
		$this->title = $string;
		return $this;
	}

	public function getRootFolder($byte) {
		return $this->root;
	}

	public function setRootFolder($byte) {
		$this->root = $byte;
		return $this;
	}

	public function getFlags() {
		return $this->flag;
	}

	public function setFlags($byte) {
		$this->flag = $byte;
		return $this;
	}
}

class FolderItem {
	private $comFolder;

	public function __construct($instance) {
		$this->comFolder = $instance;
	}

	/**
	 *
	 * @param String $property
	 *
	 * @property Application
	 * @property IsBrowsable
	 * @property IsFileSystem
	 * @property IsFolder
	 * @property IsLink
	 * @property ModifyDate
	 * @property Name
	 * @property Parent
	 * @property Path
	 * @property Size
	 * @property Type
	 *
	 *  @return multitype:boolean:string
	 */
	public function __get($property) {
		$knownProperties = array(
				"application",
				"isbrowsable",
				"isfilesystem",
				"isfolder",
				"islink",
				"modifydate",
				"name",
				"parent",
				"path",
				"size",
				"type"
		);
		if (!$this->comFolder) return NULL;
		if (array_search(strtolower($property), $knownProperties)!== false) {
			return $this->comFolder->self->$property;
		}
		throw new FolderItem_Unknown_Property_Exception("Unknown property $property");
	}

	public function getFolder() {
		if ($this->IsFolder()) {
			return $this->comFolder->self->GetFolder;
		}
		return NULL;
	}

	public function getLink() {
		if ($this->Islink()) {
			return $this->comFolder->self->GetLink;
		}
		return NULL;
	}
}