using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Nito.AsyncEx;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
namespace OutlookTest
{
    class Program
    {
        private const string Resource = "https://graph.microsoft.com/";
        private const string ResourceToken = "https://login.microsoftonline.com/";
        private const string ClientId = "";
        private const string Secret = "";
        private static readonly string Tenant = "";
        private static readonly HttpProvider HttpProvider = new HttpProvider(new HttpClientHandler(), false);
        private static readonly AuthenticationContext AuthContext = GetAuthenticationContext(Tenant);
        private static string email = "";
        private static DateTime desde = new DateTime();
        private static DateTime hasta = new DateTime();
        private static Boolean imprimiotoken;
        private static string token = "";
        public static void Main(string[] args)
        {

            try
            {
                AsyncContext.Run(() => MainAsync(args));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }

        }
        static GraphServiceClient _graphClient;
        private readonly Dictionary<string, object> _configuration;
        private static async Task MainAsync(string[] args)
        {
            while (true)
            {
                email = "";
                desde = new DateTime();
                hasta = new DateTime();
                imprimiotoken = false;
                while (true)
                {
                    var ac = 0;

                    if (email == "")
                    {
                        ac++;
                        Console.WriteLine("Agregar el email del usuario");
                        Console.ResetColor();
                        email = Console.ReadLine();

                        var spli = email.Split(',').ToList();
                        if (spli.Count == 3)
                        {

                            email = spli[0];
                            DateTime.TryParse(spli[1], out desde);
                            DateTime.TryParse(spli[2], out hasta);

                        }
                        if (email == "")
                        {
                            Colorear();
                            continue;
                        }

                    }

                    if (desde == DateTime.MinValue)
                    {
                        ac++;
                        Console.WriteLine("Agregar la fecha de inicio");
                        Console.ResetColor();
                        DateTime.TryParse(Console.ReadLine(), out desde);
                        if (desde == DateTime.MinValue)
                        {
                            Colorear();

                            continue;
                        }

                    }

                    if (hasta == DateTime.MinValue)
                    {
                        ac++;
                        Console.WriteLine("Agregar la fecha  fin");
                        Console.ResetColor();
                        DateTime.TryParse(Console.ReadLine(), out hasta);
                        if (hasta == DateTime.MinValue)
                        {
                            Colorear();

                            continue;
                        }
                    }

                    if (ac == 0)
                    {
                        break;
                    }

                }

                var lista = (await GetUserTasks(email, desde, hasta)).ToList();
                if (lista != null && lista.Count > 0)
                {

                    Console.WriteLine("------------------------------------");
                    Console.WriteLine("|        Eventos de usuario         |");
                    Console.WriteLine("------------------------------------");
                    foreach (var li in lista)
                    {
                        Console.WriteLine("El usuario {0} tiene la siguiente informacion {1}", li.UserId, li.Title);
                    }
                }
                else
                {
                    Colorear();
                    Console.WriteLine("------------------------------------");
                    Console.Write("      Sin datos del  usuario");
                    Console.ResetColor();

                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.Write(" {0} ", email);
                    Colorear();
                    Console.WriteLine("");
                    Console.WriteLine("------------------------------------");
                    Console.ResetColor();
                    Console.WriteLine("");

                }
                Console.WriteLine("------------------------------------");
                Console.WriteLine("| ¿Probar Con otro usuario? Y/N S/N |");
                Console.WriteLine("------------------------------------");
                var Salir = Console.ReadLine();
                if (Salir.ToLower() == "n" || Salir.ToLower() == "no")
                {
                    break;
                }
            }
        }

        private static void Colorear()
        {
            Console.BackgroundColor = ConsoleColor.Red;
            Console.ForegroundColor = ConsoleColor.White;
        }

        private static AuthenticationContext GetAuthenticationContext(string tenant = null)
        {
            var authString = tenant == null ?
                $"{ResourceToken}common/oauth2/token" :
                $"{ResourceToken}{tenant}/oauth2/token";

            return new AuthenticationContext(authString);
        }

        public static async Task CreateAppointment(EventModel appointment)
        {

            var authContext = new AuthenticationContext(ResourceToken + Tenant + "/oauth2/token");
            ClientCredential creds = null;
            creds = new ClientCredential(ClientId, Secret);
            var authResult = await authContext.AcquireTokenAsync(Resource, creds);
            var client = new RestClient(Resource);
            var request = new RestRequest("me/events", Method.POST);
            request.AddHeader("Authorization", "bearer " + authResult.AccessToken.ToString());
            var jsonBody = JsonConvert.SerializeObject(appointment);
            request.AddParameter("application/json", jsonBody, ParameterType.RequestBody);
            var response = client.Execute(request);

        }
        private static GraphServiceClient GetGraphClient()
        {
            var scopes = new string[] { "User.Read", "Mail.Read", "Calendars.ReadWrite", "Calendars.Read", "UserActivity.ReadWrite.CreatedByApp", "UserTimelineActivity.Write.CreatedByApp", "User.Read.All" };
            var creds = new ClientCredential(ClientId, Secret);
            AuthenticationResult result = null;
            var delegateAuthProvider = new DelegateAuthenticationProvider(async (requestMessage) =>
            {
                result = AuthContext.TokenCache?.Count > 0 ?
                              await AuthContext.AcquireTokenSilentAsync(Resource, creds,
                   new UserIdentifier(ClientId, UserIdentifierType.UniqueId)) :
                   await AuthContext.AcquireTokenAsync(Resource, creds);


                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", Token(result.AccessToken));
                //requestMessage.Headers.Add("Prefer", "outlook.body-content-type=text");
                //requestMessage.Headers.Add("Prefer", "outlook.timezone=UTC");
            });

            return new GraphServiceClient(delegateAuthProvider, HttpProvider);
        }
        static string Token(string tk)
        {
            if (!imprimiotoken)
            {
                Console.WriteLine("------------------------------------");
                Console.WriteLine("|             Token                 |");
                Console.WriteLine("------------------------------------");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(tk);
                Console.ResetColor();
                imprimiotoken = true;
            }
            token = tk;
            return tk;
        }
        private static GraphServiceClient GetGraphClient(string token)
        {
            if (token != "")
            {
                if (_graphClient != null) return _graphClient;
                _graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                            requestMessage.Headers.Add("Prefer", "outlook.body-content-type=text");
                            requestMessage.Headers.Add("Prefer", "outlook.timezone=UTC");

                            return Task.FromResult(0);
                        }));

                return _graphClient;
            }
            else
            {

                return GetGraphClient();
            }
        }
        public static async Task<IEnumerable<TaskTarea>> GetUserTasks(string mail, DateTime from, DateTime to)
        {
            var err = "";
            IUserEventsCollectionPage tasks = null;
            try
            {

                try
                {

                    var tasks1 = await GetGraphClient(token).Users[mail].Request().GetAsync();


                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine("-------------------------------------------------------");
                    Console.WriteLine("Acceso Correcto para el Usuario");
                    Console.WriteLine("         " + tasks1.Id + " " + tasks1.DisplayName + " " + tasks1.GivenName + " " + tasks1.JobTitle + "           ");
                    Console.WriteLine("         " + tasks1.City + " " + tasks1.Department + " " + tasks1.Country);
                    Console.WriteLine("         " + tasks1.MobilePhone + " " + tasks1.MySite);
                    Console.WriteLine("         " + tasks1.OfficeLocation + " " + tasks1.StreetAddress);
                    foreach (var phone in tasks1.BusinessPhones)
                    {
                        Console.WriteLine("telefono : " + phone.ToString());

                    }
                    Console.WriteLine("-------------------------------------------------------");
                    Console.ResetColor();
                    var utcFrom = from.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm");
                    var utcTo = to.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm");
                    tasks = await GetGraphClient(token).Users[mail].Events
                   .Request().Top(1000)
                   .Select("subject, body, start, end")
                   .Filter($"start/dateTime le '{utcTo}' and end/dateTime ge '{utcFrom}'")
                   .GetAsync();

                }
                catch (ServiceException e)
                {
                    switch (e.Error.Code)
                    {
                        case "Request_ResourceNotFound":
                        case "ResourceNotFound":
                        case "ErrorItemNotFound":
                        case "itemNotFound":
                            err = JsonConvert.SerializeObject(new { Message = $"User '{email}' was not found." }, Formatting.Indented);
                            break;
                        case "ErrorInvalidUser":
                            err = JsonConvert.SerializeObject(new { Message = $"The requested user '{email}' is invalid." }, Formatting.Indented);
                            break;
                        case "AuthenticationFailure":
                            err = JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
                            break;
                        //case "TokenNotFound":
                        //    await httpContext.ChallengeAsync();
                        //    err = JsonConvert.SerializeObject(new { e.Error.Message }, Formatting.Indented);
                        //    break;
                        default:
                            err = JsonConvert.SerializeObject(new { Message = e.Error.Code }, Formatting.Indented);
                            break;
                    }
                }

                return tasks.Select(x => new TaskTarea
                {
                    Title = x.Subject,
                    Status = 2,
                    Id = -1,
                    Notes = x.Body.Content,
                    Start = DateTime.Parse(x.Start.DateTime).ToUniversalTime(),
                    End = DateTime.Parse(x.End.DateTime).ToUniversalTime(),
                    Type = 4,
                    IsEditable = false,
                    ExternalId = x.Id
                });


            }
            catch (Exception e)
            {
                Colorear();
                if (err == "")
                    Console.WriteLine(e);
                else
                    Console.WriteLine(err);
                Console.ResetColor();
                return new List<TaskTarea>();
            }

        }

        public static async Task<string> Add(string mail, TaskTarea task)
        {
            var graphServiceClient = GetGraphClient();

            var outlookTask = new Event()
            {
                Subject = task.Title,
                Body = new ItemBody()
                {
                    Content = task.Notes,
                    ContentType = BodyType.Text
                },
                Start = new DateTimeTimeZone()
                {
                    DateTime = task.Start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm"),
                    TimeZone = "UTC"
                },
                End = new DateTimeTimeZone()
                {
                    DateTime = task.End.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm"),
                    TimeZone = "UTC"
                }
            };

            try
            {
                var createdTask = await graphServiceClient.Users[mail].Events.Request().AddAsync(outlookTask);
                return createdTask.Id;
            }
            catch (Exception e)
            {
                return null;
            }
        }
    }

}
