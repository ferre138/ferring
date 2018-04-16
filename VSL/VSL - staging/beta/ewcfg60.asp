<%

' ASPMaker 6 configuration file
' - contains all web site configuration settings

Const EW_PROJECT_NAME = "vsl" ' Project Name

' Session names
Dim EW_SESSION_STATUS
EW_SESSION_STATUS = EW_PROJECT_NAME & "_Status" ' Login Status
Dim EW_SESSION_USER_NAME
EW_SESSION_USER_NAME = EW_SESSION_STATUS & "_UserName" ' User Name
Dim EW_SESSION_USER_ID
EW_SESSION_USER_ID = EW_SESSION_STATUS & "_UserID" ' User ID
Dim EW_SESSION_USER_LEVEL_ID
EW_SESSION_USER_LEVEL_ID = EW_SESSION_STATUS & "_UserLevel" ' User Level ID
Dim EW_SESSION_USER_LEVEL
EW_SESSION_USER_LEVEL = EW_SESSION_STATUS & "_UserLevelValue" ' User Level
Dim EW_SESSION_PARENT_USER_ID
EW_SESSION_PARENT_USER_ID = EW_SESSION_STATUS & "_ParentUserID" ' Parent User ID
Dim EW_SESSION_SYS_ADMIN
EW_SESSION_SYS_ADMIN = EW_PROJECT_NAME & "_SysAdmin" ' System Admin
Dim EW_SESSION_AR_USER_LEVEL
EW_SESSION_AR_USER_LEVEL = EW_PROJECT_NAME & "_arUserLevel" ' User Level Array
Dim EW_SESSION_AR_USER_LEVEL_PRIV
EW_SESSION_AR_USER_LEVEL_PRIV = EW_PROJECT_NAME & "_arUserLevelPriv" ' User Level Privilege Array
Dim EW_SESSION_SECURITY
EW_SESSION_SECURITY = EW_PROJECT_NAME & "_Security" ' Security Array
Dim EW_SESSION_MESSAGE
EW_SESSION_MESSAGE = EW_PROJECT_NAME & "_Message" ' System Message

' Database settings
Dim EW_DB_CONNECTION_STRING ' DB Connection String
EW_DB_CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Request.ServerVariables("APPL_PHYSICAL_PATH") & "\db\vsldbpp.mdb" & ";"
Const EW_IS_MSACCESS = True ' Access
Const EW_IS_MSSQL = False ' MS SQL
Const EW_IS_MYSQL = False ' MySQL
Const EW_IS_ORACLE = False ' Oracle
Const EW_CURSORLOCATION = 2 ' Cursor location
Const EW_DATATYPE_NUMBER = 1
Const EW_DATATYPE_DATE = 2
Const EW_DATATYPE_TIME = 7
Const EW_DATATYPE_STRING = 3
Const EW_DATATYPE_BOOLEAN = 4
Const EW_DATATYPE_GUID = 5
Const EW_DATATYPE_OTHER = 6
Const EW_COMPOSITE_KEY_SEPARATOR = "," ' Composite key separator
Const EW_HIGHLIGHT_COMPARE = 1 ' Highlight compare mode
Const EW_ROWTYPE_VIEW = 1 ' Row type view
Const EW_ROWTYPE_ADD = 2 ' Row type add
Const EW_ROWTYPE_EDIT = 3 ' Row type edit
Const EW_ROWTYPE_SEARCH = 4 ' Row type search

' Table specific names
Const EW_TABLE_REC_PER_PAGE = "RecPerPage" ' Records per page
Const EW_TABLE_START_REC = "start" ' Start record
Const EW_TABLE_PAGE_NO = "pageno" ' Page number
Const EW_TABLE_BASIC_SEARCH = "psearch" ' Basic search keyword
Const EW_TABLE_BASIC_SEARCH_TYPE = "psearchtype" ' Basic search type
Const EW_TABLE_ADVANCED_SEARCH = "advsrch" ' Advanced search
Const EW_TABLE_SEARCH_WHERE = "searchwhere" ' Search where clause
Const EW_TABLE_WHERE = "where" ' Table where
Const EW_TABLE_ORDER_BY = "orderby" ' Table order by
Const EW_TABLE_SORT = "sort" ' Table sort
Const EW_TABLE_KEY = "key" ' Table key
Const EW_TABLE_SHOW_MASTER = "showmaster" ' Table show master
Const EW_TABLE_MASTER_TABLE = "MasterTable" ' Master table
Const EW_TABLE_MASTER_FILTER = "MasterFilter" ' Master filter
Const EW_TABLE_DETAIL_FILTER = "DetailFilter" ' Detail filter
Const EW_TABLE_RETURN_URL = "return" ' Return url

' Security specific
Const EW_AUDIT_TRAIL_PATH = "" ' Audit trail path
Const EW_ADMIN_USER_NAME = "Admin" ' Administrator user name
Const EW_ADMIN_PASSWORD = "Admin" ' Administrator password
Const EW_PARENT_USER_ID_SQL = "SELECT [CustomerID] FROM [Customers] WHERE [CustomerID] = @ParentUserID@"

' User level constants
Const EW_USER_LEVEL_COMPAT = True ' Use old user level values

'Const EW_USER_LEVEL_COMPAT = False ' Use new user level values (separate values for View/Search)
Const EW_ALLOW_ADD = 1 ' Add
Const EW_ALLOW_DELETE = 2 ' Delete
Const EW_ALLOW_EDIT = 4 ' Edit
Const EW_ALLOW_LIST = 8 ' List
Dim EW_ALLOW_VIEW, EW_ALLOW_SEARCH ' View / Search
If EW_USER_LEVEL_COMPAT Then
	EW_ALLOW_VIEW = 8 ' View
	EW_ALLOW_SEARCH = 8 ' Search
Else
	EW_ALLOW_VIEW = 32 ' View
	EW_ALLOW_SEARCH = 64 ' Search
End If
Const EW_ALLOW_REPORT = 8 ' Report
Const EW_ALLOW_ADMIN = 16 ' Admin

' Date separator
Const EW_DATE_SEPARATOR = "/"

' Email related constants
Const EW_SMTP_SERVER = "smtp.1and1.com" ' Smtp server
Const EW_SMTP_SERVER_PORT = 25 ' Smtp server port
Const EW_SMTP_SERVER_USERNAME = "info@vsl3.ca" ' Smtp server user name
Const EW_SMTP_SERVER_PASSWORD = "web4colon" ' Smtp server password
Const EW_SENDER_EMAIL = "info@vsl3.ca" ' Sender email
Const EW_RECIPIENT_EMAIL = "" ' Receiver email
' File upload constants
Const EW_UPLOAD_DEST_PATH = "" ' Upload destination path
Const EW_UPLOAD_ALLOWED_FILE_EXT = "gif,jpg,jpeg,bmp,png,doc,xls,pdf,zip" ' Allowed file extensions
Const EW_UPLOAD_CHARSET = "" ' Upload charset
Const EW_MAX_FILE_SIZE = 2000000 ' Max file size
Const EW_THUMBNAIL_FILE_PREFIX = "tn_" ' Thumbnail file prefix
Const EW_THUMBNAIL_FILE_SUFFIX = "" ' Thumbnail file suffix
Const EW_THUMBNAIL_DEFAULT_WIDTH = 0 ' Thumbnail default width
Const EW_THUMBNAIL_DEFAULT_HEIGHT = 0 ' Thumbnail default height
Const EW_THUMBNAIL_DEFAULT_INTERPOLATION = 1 ' Thumbnail default interpolation

' Export all records
Const EW_EXPORT_ALL = True ' export all records

' Const EW_EXPORT_ALL = False ' export 1 page only
%>
