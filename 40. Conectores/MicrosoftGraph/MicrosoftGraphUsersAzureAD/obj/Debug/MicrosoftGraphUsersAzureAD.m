 /*******************************************************************************************
 *																                            *
 *	MicrosoftGraphUsersAzureAD.pq      							                            *
 * 																                            *
 *	Creado por: 	Carlos Pesquera Nieto						                            *
 *	Fecha creación:	20/04/2020                                                              *
 *	Descripción: 	Fichero que contiene la lógica del conector Power Query para usar en    *
 *                  Power BI y conectarse y consultar los usuarios del Azure Active         *
 *                  Directory.                                                              * 
 *																                            *
 * Documentación:   ¿Qué es Microsoft Graph?                                                *
 *                  Microsoft Graph es una API que contiene gran cantidad de datos          *
 *                  disponibles en Office 365, Enterprise Mobility + Security y Windows 10. *
 *                  Tambien nos permite acceso a  Azure AD, Excel, Intune, Outlook/Exchange,* 
 *                  OneDrive, OneNote, SharePoint, Planner y muchos más mediante un único   *
 *                  punto de conexión.                                                      *
 *                  https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0 *
 *                                                                                          *
 ********************************************************************************************/

section MicrosoftGraphUsersAzureAD;

//
// Configuración de las variables para OAuth
//

// Leemos y establecemos el valor del ID de la aplicación (Application ID) previamente registrada en portal.azure.com
client_id = Text.FromBinary(Extension.Contents("client_id"));

// Leemos y establecemos el valor del secreto de la aplicación (Application secret) previamente registrada en portal.azure.com
// NOTA: Este valor de Application secret, hay que copiarselo al registrar la aplicación por primera vez porque luego se oculta.
client_secret = Text.FromBinary(Extension.Contents("client_secret"));

redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html";
token_uri = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token";
authorize_uri = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize";
logout_uri = "https://login.microsoftonline.com/logout.srf";

// Es necesario el alcance "offline_access" para recibir el valor de la actualización del token.
// Esto se agrega por separado desde los ámbitos de Microsoft Graph.
// Ver https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#offline_access
//
// Para obtener más información de los ambitos disponibles en Microsoft Graph:
//      - Información general desde:
//          https://developer.microsoft.com/en-us/graph/docs/authorization/permission_scopes
//      - Ambitos necesarios para obtener los usuarios de Azure AD, sección "Authorization" desde:
//          https://docs.microsoft.com/en-us/graph/api/resources/users?view=graph-rest-1.0
scope_prefix = "https://graph.microsoft.com/";

// Ambito de los permisos requeridos que deben ser configurados en la aplicación registrada para Microsoft Graph en portal.azure.com sección, "Permisos de API"
scopes = {
    "User.ReadBasic.All",
    "User.Read",
    "User.ReadWrite",
    "User.Read.All",
    "User.ReadWrite.All",
    "Directory.Read.All",
    "Directory.ReadWrite.All",
    "Directory.AccessAsUser.All"
};

//
// Definición de las funciones de llamada
//
[DataSource.Kind="MicrosoftGraphUsersAzureAD", Publish="MicrosoftGraphUsersAzureAD.UI"]
shared MicrosoftGraphUsersAzureAD.Feed = () =>
    let
        Origen = OData.Feed("https://graph.microsoft.com/beta/users/", null, [ ODataVersion = 4, MoreColumns = true ])
    in
        Origen;

//
// Definición del Datasource para OAuth
// Las funciones StartLogin, FinishLogin, Refresh y Logout están definidas más abajo: Implementación OAuth
//
MicrosoftGraphUsersAzureAD = [
    TestConnection = (dataSourcePath) => { "MicrosoftGraphUsersAzureAD.Feed" },
    Authentication = [
        OAuth = [
            StartLogin=StartLogin,
            FinishLogin=FinishLogin,
            Refresh=Refresh,
            Logout=Logout
        ]
    ],
    Label = "Conector Microsoft Graph para la obtención de usuarios desde Azure Active Directory"
];

//
// Definición de la interfaz UI del Datasource
//
MicrosoftGraphUsersAzureAD.UI = [
    Beta = true,
    ButtonText = { "MicrosoftGraphUsersAzureAD.Feed", "Connectar a Microsoft Graph" },
    SourceImage = MicrosoftGraphUsersAzureAD.Icons,
    SourceTypeImage = MicrosoftGraphUsersAzureAD.Icons
];

MicrosoftGraphUsersAzureAD.Icons = [
    Icon16 = { Extension.Contents("MicrosoftGraphUsersAzureAD16.png"), Extension.Contents("MicrosoftGraphUsersAzureAD20.png"), Extension.Contents("MicrosoftGraphUsersAzureAD24.png"), Extension.Contents("MicrosoftGraphUsersAzureAD32.png") },
    Icon32 = { Extension.Contents("MicrosoftGraphUsersAzureAD32.png"), Extension.Contents("MicrosoftGraphUsersAzureAD40.png"), Extension.Contents("MicrosoftGraphUsersAzureAD48.png"), Extension.Contents("MicrosoftGraphUsersAzureAD64.png") }
];

//
// Implementación funciones OAuth
//
// Puedes ver los siguientes enlaces para más información y detalles del funcionamiento sobre AAD/Graph OAuth:
//      https://docs.microsoft.com/en-us/azure/active-directory/active-directory-protocols-oauth-code 
//
// La funcion StartLogin crea un registro que contiene la información necesaria para que el cliente de OAuth inicie un flujo.
// NOTA: para el flujo de Azure AD, el parámetro de visualización no se utiliza.
//
// resourceUrl: Se deriva de los argumentos requeridos para la función del Datasource y se usa cuando el flujo OAuth 
//              requiere que se pase un recurso específico o se calcula la URL de autorización (es decir, cuando el 
//              nombre/ID del Tenant se incluye en la URL). Aquí estamos hardcodeando el uso del Tenant "común", según
//              se lo estamos especificando en la variable 'authorize_uri'.
// state:       Valor del estado del cliente que le pasamos al servicio.
// display:     Utilizado por ciertos servicios de OAuth para mostrar información al usuario
//
// Esta función DEVUELVE un registro que contiene los siguientes campos:
// LoginUri:     La URI completa a usar cuando se inicia el cuadro de dialogo del flujo OAuth.
// CallbackUri:  El valor de redirect_uri. El cliente considerará el flujo completo de OAuth
//               cuando reciba una redirección a esta URI. Esto genearalmente debe coincidir
//               con el valor de redirect_uri que se registró para nuestra aplicacion/cliente.
// WindowHeight: Altura de ventana sugerida para OAuth (en pixels).
// WindowWidth:  Ancho de ventana sugerido para OAuth (en pixels).
// Context:      Valor de contexto opcional que se le pasara a la función FinishLogin una vez
//               que es alcanzada la redirect_uri.
//

StartLogin = (resourceUrl, state, display) =>
    let
        authorizeUrl = authorize_uri & "?" & Uri.BuildQueryString([
            client_id = client_id,  
            redirect_uri = redirect_uri,
            state = state,
            scope = "offline_access " & GetScopeString(scopes, scope_prefix),
            response_type = "code",
            response_mode = "query",
            login = "login"
        ])
    in
        [
            LoginUri = authorizeUrl,
            CallbackUri = redirect_uri,
            WindowHeight = 720,
            WindowWidth = 1024,
            Context = null
        ];

//
// La función FinishLogin se llama cuando el flujo OAuth alcanza la redirect_uri especificada.
// NOTA: para el flujo de Azure AD, no se utilizan los parametros de context y state.
//
// context:     El valor para el campo Context devuelto por StartLogin. Esto se usa cuando deseemos 
//              pasar informacion derivada durante la llamada a la función StartLogin call (como el 
//              ID del Tenant)
// callbackUri: El callbackUri contiene el authorization_code (código de autorizacion) del servicio.
// state:       Información del estado que se especificó durante la llamada de la función StartLogin.
//
FinishLogin = (context, callbackUri, state) =>
    let
        // parseamos completamente la variable callbackUri y extraemos la Query
        parts = Uri.Parts(callbackUri)[Query],
        // si la cadena de la Query contiene un campo de error, lanzamos un error
        // de lo contrario, llamamos a la funcion TokenMethod para intercambiar nuestro access_token (código de acceso)
        result = if (Record.HasFields(parts, {"error", "error_description"})) then 
                    error Error.Record(parts[error], parts[error_description], parts)
                 else
                    TokenMethod("authorization_code", "code", parts[code])
    in
        result;

//
// Funcion que llama cuando el access_token (codigo de acceso) ha caducado y hay un refresh_token (refresco del token) disponible.
// 
Refresh = (resourceUrl, refresh_token) => TokenMethod("refresh_token", "refresh_token", refresh_token);

//
// Funcion de cierre de sesion
//
Logout = (token) => logout_uri;

//
// Función TokenMethod para intercambiar nuestro access_token (código de acceso)
//
// grantType:  Se mapea el parametro "grant_type" a la query.
// tokenField: El nombre del parametro de la query que se pasa en code.
// code:       Es el codigo "bueno" (authorization_code o refresh_token) que se envia al servicio.
//
TokenMethod = (grantType, tokenField, code) =>
    let
        queryString = [
            client_id = client_id,
            client_secret = client_secret,
            scope = "offline_access " & GetScopeString(scopes, scope_prefix),
            grant_type = grantType,
            redirect_uri = redirect_uri
        ],
        queryWithCode = Record.AddField(queryString, tokenField, code),

        tokenResponse = Web.Contents(token_uri, [
            Content = Text.ToBinary(Uri.BuildQueryString(queryWithCode)),
            Headers = [
                #"Content-type" = "application/x-www-form-urlencoded",
                #"Accept" = "application/json"
            ],
            ManualStatusHandling = {400} 
        ]),
        body = Json.Document(tokenResponse),
        result = if (Record.HasFields(body, {"error", "error_description"})) then 
                    error Error.Record(body[error], body[error_description], body)
                 else
                    body
    in
        result;

//
// Funciones auxiliares
//
Value.IfNull = (a, b) => if a <> null then a else b;

GetScopeString = (scopes as list, optional scopePrefix as text) as text =>
    let
        prefix = Value.IfNull(scopePrefix, ""),
        addPrefix = List.Transform(scopes, each prefix & _),
        asText = Text.Combine(addPrefix, " ")
    in
        asText;
