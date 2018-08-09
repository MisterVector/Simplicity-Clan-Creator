Attribute VB_Name = "modVars"
Public Const PROGRAM_VERSION                  As String = "1.3.2"
Public Const PROGRAM_TITLE                    As String = "Simplicity v" & PROGRAM_VERSION & " by Vector"

Public Const RELEASES_URL                     As String = "https://github.com/MisterVector/Simpilcity-Clan-Creator-Legacy/releases"

Public Const DEFAULT_SERVER                   As String = "useast.battle.net"
Public Const DEFAULT_BNLS_SERVER              As String = "jbls.codespeak.org"
Public Const DEFAULT_CHANNEL                  As String = "Simplicity"
Public Const DEFAULT_CONNECTIONS_PER_PROXY    As Integer = 4
Public Const DEFAULT_TIMEOUT                  As Integer = 10000
Public Const DEFAULT_RECONNECT_TIME           As Integer = 12000
Public Const DEFAULT_VERSION_BYTE             As Long = &H1E

Public Const DEFAULT_REMEMBER_WINDOW_POSITION As Boolean = False
Public Const DEFAULT_SAVE_CLAN_INFO           As Boolean = True
Public Const DEFAULT_USE_CUSTOM_INITIATES     As Boolean = False
Public Const DEFAULT_CHECK_UPDATE_ON_STARTUP  As Boolean = True

Public initiateManager As New clsCustomInitiates
Public initiateNumber As Integer

Public botsAdded As New Dictionary

Public dicServerList As New Dictionary

Public dicInitiatesAdded As New Dictionary

Public oldInitiateName As String
Public oldChieftainName As String

Public updateString As String
Public connectedCount As Integer

Public Hashes(0) As String
Public isCheckingClanTag As Boolean
Public Packet() As clsPacketHandler
Public bnlsPacketHandler As clsPacketHandler
Public chiefPacketHandler As clsPacketHandler
Public programLoaded As Boolean
Public isCreatingClan As Boolean
Public attemptedVerByteUpdate As Boolean

Public keys As New clsKeyManager
Public chiefData As New clsChieftainData

Public clannedKeyCheckClan As String

Public manualUpdateCheck As Boolean

Public Enum KeyType
    BAD
    IN_USE
    CLANNED
End Enum

Public Enum packetType
    BNCS
    BNLS
End Enum

Public Type ConfigType
    server As String
    bnlsServer As String
    timeOut As Long
    reconnectTime As Integer
    verByte As Long
    initiate As String
    connectsPerProxy As Integer
    initiatePassword As String
    Channel As String
    clanTag As String
    clanDescription As String
    saveClanInfo As Boolean
    useCustomInitiates As Boolean
    
    checkUpdateOnStartup As Boolean
    rememberWindowPosition As Boolean
    windowTop As Long
    windowLeft As Long
End Type
Public config As ConfigType

Public Type ClientData
    username As String
    password As String
    key As String
    proxyVersion As String
    keyIndex As String
    hasRestrictedKey As Boolean
    
    hasChieftainAsFriend As Boolean
    hasCheckedFriendsList As Boolean
    
    proxyIP As String
    proxyPort As Long
    
    nls_P As Long
    
    isReadyForPreparation As Boolean
    hasDoneProcedure As Boolean
    loggedOn As Boolean
    hasCheckedKey As Boolean
End Type
Public bot() As ClientData
Public chief As ClientData

