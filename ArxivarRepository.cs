using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using ACUtils.AXRepository.Exceptions;


namespace ACUtils.AXRepository
{
    public class ArxivarRepository
    {
        const string STATO_ELIMINATO = "ELIMINATO";
        protected string _username;
        protected string _password;
        protected string _apiUrl;
        protected string _managementUrl;
        protected string _workflowUrl;
        protected string _appId;
        protected string _appSecret;
        protected string _aoo;
        protected string _wcfUrl;
        protected long? _impersonateUserId;

        protected string _token;
        protected string _refreshToken;

        protected string _tokenManagement;
        protected string _refreshTokenManagement;

        ACUtils.ILogger? _logger = null;

        #region configuration

        protected Abletech.WebApi.Client.Arxivar.Client.Configuration configuration =>
            new Abletech.WebApi.Client.Arxivar.Client.Configuration()
            {
                ApiKey = new Dictionary<string, string>() { { "Authorization", _token } },
                ApiKeyPrefix = new Dictionary<string, string>() { { "Authorization", "Bearer" } },
                BasePath = _apiUrl,
            };

        protected Abletech.WebApi.Client.ArxivarManagement.Client.Configuration configurationManagement =>
            new Abletech.WebApi.Client.ArxivarManagement.Client.Configuration()
            {
                ApiKey = new Dictionary<string, string>() { { "Authorization", _tokenManagement } },
                ApiKeyPrefix = new Dictionary<string, string>() { { "Authorization", "Bearer" } },
                BasePath = _managementUrl,
            };

        protected Abletech.WebApi.Client.ArxivarWorkflow.Client.Configuration configurationWorkflow =>
            new Abletech.WebApi.Client.ArxivarWorkflow.Client.Configuration()
            {
                ApiKey = new Dictionary<string, string>() { { "Authorization", _token } },
                ApiKeyPrefix = new Dictionary<string, string>() { { "Authorization", "Bearer" } },
                BasePath = _workflowUrl,
            };

        #endregion

        #region constructor

        public ArxivarRepository(
            string apiUrl, string managementUrl, string workflowUrl, string username, string password, string appId,
            string appSecret, string AOO,
            string wcf_url = "net.tcp://127.0.0.1:8740/Arxivar/Push",
            ACUtils.ILogger? logger = null,
            long? impersonateUserId = null
        )
        {
            this._apiUrl = apiUrl;
            this._managementUrl = managementUrl;
            this._workflowUrl = workflowUrl;
            this._username = username;
            this._password = password;
            this._appId = appId;
            this._appSecret = appSecret;
            this._logger = logger;
            this._wcfUrl = wcf_url;
            this._aoo = AOO;
            this._impersonateUserId = impersonateUserId;
        }


        public ArxivarRepository(
            string apiUrl,
            string managementUrl,
            string workflowUrl,
            string authToken,
            string AO,
            ACUtils.ILogger logger = null
        )
        {
            this._apiUrl = apiUrl;
            this._managementUrl = managementUrl;
            this._workflowUrl = workflowUrl;
            this._logger = logger;
            this._aoo = AO;
            _token = authToken;
        }


        public void authToken(string token, string refreshToken)
        {
            this._token = token;
            this._refreshToken = refreshToken;
        }

        #endregion

        #region file upload

        public List<string> UploadFile(Stream stream, bool cacheInsert = false)
        {
            Login();
            if (cacheInsert)
            {
                var cacheApi = new Abletech.WebApi.Client.Arxivar.Api.CacheApi(configuration);
                return cacheApi.CacheInsert(stream);
            }
            else
            {
                var bufferApi = new Abletech.WebApi.Client.Arxivar.Api.BufferApi(configuration);
                return bufferApi.BufferInsert(stream);
            }
        }

        private string GetTemporaryDirectory()
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }

        public List<string> UploadFile(string filename, byte[] bytes, bool cacheInsert = false)
        {
            Login();
            
            var tmpDir = GetTemporaryDirectory();
            try
            {
                var tmppath = Path.Combine(tmpDir, filename);
                using (var stream = new MemoryStream(bytes))
                {
                    using (var fileStream = new FileStream(tmppath, FileMode.CreateNew, FileAccess.ReadWrite))
                    {
                        stream.WriteTo(fileStream);
                        fileStream.Flush();
                        fileStream.Seek(0, SeekOrigin.Begin);
                        return UploadFile(fileStream, cacheInsert);
                    }
                }
            }
            finally
            {
                try
                {
                    System.IO.Directory.Delete(tmpDir, true);
                }
                catch
                {
                }
            }
        }

        public List<string> UploadFile(string filePath, string filename = null, bool cacheInsert = false)
        {
            Login();

            var bufferApi = new Abletech.WebApi.Client.Arxivar.Api.BufferApi(configuration);

            // workaround per aprire file su share di rete e cambiare nome al file caricato
            var tmpDir = GetTemporaryDirectory();
            try
            {
                var tmppath = Path.Combine(tmpDir, filename ?? Path.GetFileName(filePath));
                File.Copy(filePath, tmppath);

                using (var stream = File.Open(tmppath, FileMode.Open))
                {
                    return UploadFile(stream, cacheInsert);
                }
            }
            finally
            {
                try
                {
                    System.IO.Directory.Delete(tmpDir, true);
                }
                catch
                {
                }
            }
        }

        #endregion


        #region address book

        /// <summary>
        /// 
        /// </summary>
        /// <param name="codice"></param>
        /// <param name="addressBookCategoryId"></param>
        /// <param name="type">Possible values:  To => 0 | From => 1 |  CC => 2 | Senders => 3</param>
        /// <returns></returns>
        public Abletech.WebApi.Client.Arxivar.Model.UserProfileDTO GetAddressBookEntry(string codice,
            int addressBookCategoryId, UserProfileType type = UserProfileType.To)
        {
            Login();

            var addressBookApi = new Abletech.WebApi.Client.Arxivar.Api.AddressBookApi(configuration);
            var filter = addressBookApi.AddressBookGetSearchField();
            var select = addressBookApi.AddressBookGetSelectField()
                .Select("DM_RUBRICA_CODICE")
                .Select("DM_RUBRICA_AOO")
                .Select("DM_RUBRICA_CODICE")
                .Select("DM_RUBRICA_SYSTEM_ID")
                .Select("ID");

            var results = addressBookApi.AddressBookPostSearch(
                new Abletech.WebApi.Client.Arxivar.Model.AddressBookSearchCriteriaDTO(
                    filter: codice,
                    addressBookCategoryId: addressBookCategoryId,
                    filterFields: filter,
                    selectFields: select
                ));

            var result = results.Data.First();

            var contactId = result.GetValue<int>("ID");  // workaround: ID è di tipo long ma le funzioni che prendono ID come parametro sono di tipo int.
            var addressBookId = result.GetValue<int>("DM_RUBRICA_SYSTEM_ID"); // workaround: come sopra, è un long ma le funzioni accettano un int.

            var addressBook = addressBookApi.AddressBookGetById(addressBookId: addressBookId);
            return new Abletech.WebApi.Client.Arxivar.Model.UserProfileDTO(
                id: addressBook.Id,
                externalId: addressBook.ExternalCode,
                description: addressBook.BusinessName,
                docNumber: "-1",
                type: (int)type,
                contactId: contactId,
                fax: addressBook.Fax,
                address: addressBook.Address,
                postalCode: addressBook.PostalCode,
                contact: "",
                job: "",
                locality: addressBook.Location,
                province: addressBook.Province,
                phone: addressBook.PhoneNumber,
                mobilePhone: addressBook.CellPhone,
                telName: "",
                faxName: "",
                house: "",
                department: "",
                reference: "",
                office: "",
                vat: "",
                mail: "",
                priority: "N", // addressBook.Priority,
                code: null,
                email: addressBook.Email,
                fiscalCode: addressBook.FiscalCode,
                nation: addressBook.Country,
                addressBookId: addressBook.Id,
                society: ""
                //officeCode: "",
                //publicAdministrationCode: "",
                //pecAddressBook: "",
                //feaEnabled: false,
                //feaExpireDate: null,
                //firstName: "",
                //lastName: "",
                //pec: ""
            );
        }

        #endregion

        #region auth

        private Abletech.WebApi.Client.Arxivar.Model.AuthenticationTokenDTO _login(List<string> scopeList = null)
        {
            var authApi = new Abletech.WebApi.Client.Arxivar.Api.AuthenticationApi(_apiUrl);
            var auth = authApi.AuthenticationGetToken(
                new Abletech.WebApi.Client.Arxivar.Model.AuthenticationTokenRequestDTO(
                    username: _username,
                    password: _password,
                    clientId: _appId,
                    clientSecret: _appSecret,
                    impersonateUserId: _impersonateUserId.HasValue
                        ? System.Convert.ToInt32(_impersonateUserId)
                        : default(int?),
                    scopeList: scopeList
                )
            );
            return auth;
        }


        private void Login()
        {
            if (string.IsNullOrEmpty(_token)) // TODO: test se è necessario il refresh del token
            {
                var auth = _login();
                _token = auth.AccessToken;
                _refreshToken = auth.RefreshToken;
            }
        }


        private void LoginManagment()
        {
            if (string.IsNullOrEmpty(_tokenManagement))
            {
                var scopeList = new List<string> { "ArxManagement" };
                var auth = _login(scopeList);
                _tokenManagement = auth.AccessToken;
                _refreshTokenManagement = auth.RefreshToken;
            }
        }

        #endregion

        #region profile - get

        public T GetProfile<T>(int docnumber) where T : AXModel<T>, new()
        {
            Login();
            var profileApi = new Abletech.WebApi.Client.Arxivar.Api.ProfilesApi(configuration);
            var profile = profileApi.ProfilesGetSchema(docnumber, false);
            var obj = AXModel<T>.Idrate(profile);
            return obj;
        }

        #endregion

        #region profile - search

        public List<T> Search<T>(AXModel<T> model, bool eliminato = false) where T : AXModel<T>, new()
        {
            var searchValues = model.GetPrimaryKeys();
            return Search<T>(searchValues, eliminato);
        }

        public List<T> Search<T>(Dictionary<string, object> searchValues = null, bool eliminato = false)
            where T : AXModel<T>, new()
        {
            var classeDoc = (new T()).GetArxivarAttribute().DocumentType;
            var result = RawSearch(
                classeDoc: classeDoc,
                searchValues: searchValues,
                eliminato: eliminato
            );

            var profiles = result.Select(s => GetProfile<T>(s.Columns.GetValue<int>("DOCNUMBER"))).ToList();
            return profiles;
        }


        public List<Abletech.WebApi.Client.Arxivar.Model.RowSearchResult> RawSearch(string classeDoc,
            Dictionary<string, object> searchValues = null, bool eliminato = false, bool selectAll = false)
        {
            Login();
            
            

            var searchApi = new Abletech.WebApi.Client.Arxivar.Api.SearchesApi(configuration);
            var docTypesApi = new Abletech.WebApi.Client.Arxivar.Api.DocumentTypesApi(configuration);
            var docTypes = docTypesApi.DocumentTypesGetOld(1, this._aoo); // TODO replace deprecated method

            var classeDocumento = docTypes.First(i => i.Key == classeDoc);

            var filterSearch = searchApi.SearchesGet()
                .Set("DOCUMENTTYPE",
                    new Abletech.WebApi.Client.Arxivar.Model.DocumentTypeSearchFilterDto(classeDocumento.DocumentType,
                        classeDocumento.Type2, classeDocumento.Type3));
            ;

            var defaultSelect = searchApi
                .SearchesGetSelect_1(classeDocumento.DocumentType, classeDocumento.Type2, classeDocumento.Type3)
                .Select("WORKFLOW")
                .Select("DOCNUMBER");


            var additionals = searchApi.SearchesGetAdditionalByClasse(classeDocumento.DocumentType,
                classeDocumento.Type2, classeDocumento.Type3, this._aoo);
            filterSearch.Fields.AddRange(additionals);

            if (!(searchValues is null))
            {
                foreach (var kv in searchValues)
                {
                    filterSearch.Set(kv.Key, kv.Value);
                }
            }

            if (selectAll)
            {
                foreach (var field in defaultSelect.Fields)
                {
                    if (field.FieldType == 2)
                        defaultSelect.Select(field.Name);
                }
            }


            if (!eliminato)
            {
                filterSearch.Set("Stato", STATO_ELIMINATO, 2); // diverso da ELIMINATO
            }

            var values =
                searchApi.SearchesPostSearch(
                    new Abletech.WebApi.Client.Arxivar.Model.SearchCriteriaDto(filterSearch, defaultSelect));
            return values;
        }

        public int GetDocumentNumber(string classeDoc, Dictionary<string, object> searchValues, bool eliminato = false,
            bool getFirst = false)
        {
            Login();
            var values = RawSearch(classeDoc: classeDoc, searchValues: searchValues, eliminato: eliminato);

            if (values.Count > 1)
            {
                if (getFirst)
                {
                    // con l'opzione getFirst viene restituto il documento con DOCNUMBER inferiore

                    // ordina i risultati per DOCNUMBER
                    values.Sort((x, y) =>
                    {
                        // Valori restituiti:
                        //     Intero con segno che indica i valori relativi di x e y, come illustrato nella
                        //     tabella seguente. Valore Significato Minore di zero x è minore di y. Zero x è
                        //     uguale a y. Maggiore di zero x è maggiore di y.
                        var xVal = x.GetValue<int>("DOCNUMBER");
                        var yVal = y.GetValue<int>("DOCNUMBER");
                        return xVal.CompareTo(yVal);
                    });
                }
                else
                {
                    throw new TooMuchElementsException($"La ricerca ha ricevuto {values.Count} risultati");
                }
            }

            if (values.Count == 0)
            {
                throw new NotFoundException($"La ricerca ha ricevuto nessun risultato");
            }

            var docNumber = (int)values.First().Get("DOCNUMBER").Value;
            return docNumber;
        }

        public int GetDocumentNumber<T>(AXModel<T> model, bool eliminato = false, bool getFirst = false)
            where T : AXModel<T>, new()
        {
            var searchValues = model.GetPrimaryKeys();
            var classeDoc = model.GetArxivarAttribute().DocumentType;
            return GetDocumentNumber(
                classeDoc: classeDoc,
                searchValues: searchValues,
                eliminato: eliminato,
                getFirst: getFirst
            );
        }

        #endregion

        #region profile - update

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <param name="taskid">se il documento è sottoposto a workflow è necessario passare anche il numero di task</param>
        /// <param name="procdocid">se il documento è sottoposto a workflow è necessario passare anche il numero di processo</param>
        /// <param name="checkInOption"></param>
        /// <param name="killWorkflow"></param>
        /// <returns></returns>
        public long? UpdateProfile<T>(AXModel<T> model, int? taskid = null, int? procdocid = null,
            int checkInOption = 0, bool killWorkflow = false) where T : AXModel<T>, new()
        {
            Login();
            List<string> bufferIds = new List<string>();
            T doc;
            if (model.DOCNUMBER.HasValue)
            {
                doc = GetProfile<T>(model.DOCNUMBER.Value);
            }
            else if (model.GetPrimaryKeys().Count > 0)
            {
                doc = Search<T>(model).First();
            }
            else
            {
                throw new Exception("Cannot update profile without DOCNUMBER or primary keys");
            }

            var workflow = System.Convert.ToInt64(doc.Workflow ?? false);

            //var docNumber = model.DOCNUMBER ?? GetDocumentNumber(model);
            var docNumber = model.DOCNUMBER ?? doc.DOCNUMBER;

            if (workflow == 1 && killWorkflow)
            {
                try
                {
                    var workflowApi = new Abletech.WebApi.Client.Arxivar.Api.WorkflowApi(configuration);
                    var workflowHistory =
                        workflowApi.WorkflowGetWorkflowInfoByDocnumber(System.Convert.ToInt32(docNumber));

                    var processId = workflowHistory.Where(w => w.State == 1).Select(w => w.Id).FirstOrDefault();
                    if (processId.HasValue)
                    {
                        var taskworkhistoryapi =
                            new Abletech.WebApi.Client.Arxivar.Api.TaskWorkHistoryApi(configuration);
                        var history = taskworkhistoryapi.TaskWorkHistoryGetHistoryByProcessId(processId);
                        if (!history.Where(r => !r.GetValue<DateTime?>("CONCLUSO").HasValue) .Any()) // WTF? controllare questa condizione
                        {
                            DeleteWorkflow(processId);
                        }

                        workflowApi.WorkflowFreeUserConstraint(processId.Value);
                    }

                    var processDocumentApi = new Abletech.WebApi.Client.Arxivar.Api.ProcessDocumentApi(configuration);
                    processDocumentApi.ProcessDocumentFreeWorkflowConstraint(System.Convert.ToInt32(docNumber));
                }
                catch
                {
                }
            }

            var profilesApi = new Abletech.WebApi.Client.Arxivar.Api.ProfilesApi(configuration);

            var profileDto = profilesApi.ProfilesGetSchema(docNumber, true);

            if (model.DataDoc.HasValue)
            {
                profileDto.SetField("DataDoc", model.DataDoc.Value);
            }

            if (!string.IsNullOrEmpty(model.STATO))
            {
                profileDto.SetState(model.STATO);
            }

            /*
            try
            {
                profileDto.Fields.SetFromField(GetUserAddressBookEntry(model.User, 1));
            }
            catch { }
            */

            foreach (var field in model.GetArxivarFields())
            {
                try
                {
                    profileDto.SetField(field.Key, field.Value);
                }
                catch (AXFieldNotFoundException)
                {
                    if (
                        field.Key != "From_ExternalId" &&
                        field.Key != "To_ExternalId"
                    ) throw;
                }
            }

            if (!string.IsNullOrEmpty(model.FilePath) || model.File.HasValue)
            {
                var checkInOutApi = new Abletech.WebApi.Client.Arxivar.Api.CheckInOutApi(configuration);
                //var taskWorkApi = new ArxivarNext.Api.TaskWorkApi(configuration);

                var isCheckOutForTask = taskid.HasValue && procdocid.HasValue;

                // TODO: prevedere la possibilità di checkin/checkout for task
                if (isCheckOutForTask)
                {
                    //var select = taskWorkApi.TaskWorkGetDefaultSelect();
                    //var tasks = taskWorkApi.TaskWorkGetActiveTaskWork(select, System.Convert.ToInt32(docNumber));
                    //var taskWork = taskWorkApi.TaskWorkGetTaskWorkById(taskid);
                    //checkInOutApi.CheckInOutCheckOutForTask(processDocId: procdocid, taskWorkId: taskid);
                }
                else
                {
                    checkInOutApi.CheckInOutCheckOut(System.Convert.ToInt32(docNumber));
                }

                if (model.File.HasValue)
                {
                    bufferIds = UploadFile(model.File.Value.name, model.File.Value.bytes, cacheInsert: true);
                }
                else
                {
                    bufferIds = UploadFile(model.FilePath, cacheInsert: true);
                }

                if (isCheckOutForTask)
                {
                    checkInOutApi.CheckInOutCheckInForTask(
                        processDocId: procdocid,
                        taskWorkId: taskid,
                        fileId: bufferIds.First()
                    );
                }
                else
                {
                    checkInOutApi.CheckInOutCheckIn(
                        docnumber: System.Convert.ToInt32(docNumber),
                        fileId: bufferIds.First(),
                        option: checkInOption,
                        undoCheckOut: true
                    );
                }
            }

            profilesApi.ProfilesPut(docNumber, new Abletech.WebApi.Client.Arxivar.Model.ProfileDTO()
            {
                Fields = profileDto.Fields,
                Document = bufferIds.Count > 0
                    ? new Abletech.WebApi.Client.Arxivar.Model.FileDTO() { BufferIds = bufferIds }
                    : default(Abletech.WebApi.Client.Arxivar.Model.FileDTO),
            });

            return docNumber;
        }

        #endregion

        #region profile - delete

        public bool DeleteProfile<T>(AXModel<T> model) where T : AXModel<T>, new()
        {
            Login();

            int docNumber;
            try
            {
                // vabene prendere il primo documento trovato
                // si suppone che se il documento corretto è già stato inserito questo avrà un DOCNUMBER superiore a quello da eliminare
                // l'opzione getFirst ottiene il documento con DOCNUMBER inferiore
                docNumber = GetDocumentNumber(model, getFirst: true);
            }
            catch (NotFoundException)
            {
                return true;
            }

            var profilesApi = new Abletech.WebApi.Client.Arxivar.Api.ProfilesApi(configuration);
            var profileDto = profilesApi.ProfilesGetSchema(docNumber, true);
            profileDto.SetState(STATO_ELIMINATO);
            profilesApi.ProfilesPut(docNumber, new Abletech.WebApi.Client.Arxivar.Model.ProfileDTO()
            {
                Fields = profileDto.Fields
            });
            return true;
        }

        public bool HardDeleteProfile(int docNumber)
        {
            Login();
            var profilesApi = new Abletech.WebApi.Client.Arxivar.Api.ProfilesApi(configuration);
            profilesApi.ProfilesDeleteProfile(docNumber);
            return true;
        }

        #endregion

        #region profile - create

        public int? CreateProfile<T>(AXModel<T> model, bool updateIfExists = false, int checkInOption = 0,
            bool killworkflow = false) where T : AXModel<T>, new()
        {
            Login();

            if (!model.GetArxivarAttribute().SkipKeyCheck)
            {
                var search = this.Search(model);
                if (search.Any())
                {
                    if (updateIfExists)
                    {
                        model.DOCNUMBER = search.First().DOCNUMBER; // avoid search again during update 
                        var docNumber = UpdateProfile(model, checkInOption: checkInOption, killWorkflow: killworkflow);
                        return System.Convert.ToInt32(docNumber);
                    }
                    else
                    {
                        return System.Convert.ToInt32(search.First().DOCNUMBER);
                    }
                }
            }

            var profileApi = new Abletech.WebApi.Client.Arxivar.Api.ProfilesApi(configuration);
            var statesApi = new Abletech.WebApi.Client.Arxivar.Api.StatesApi(configuration);


            List<string> bufferId = new List<string>();
            if (model.File.HasValue)
            {
                bufferId = UploadFile(model.File.Value.name, model.File.Value.bytes);
            }
            else if (!string.IsNullOrEmpty(model.FilePath))
            {
                bufferId = UploadFile(model.FilePath);
            }

            var documentType = model.GetArxivarAttribute().DocumentType;

            var profileDto = profileApi.ProfilesGet_0();
            profileDto.Attachments = new List<string>();
            profileDto.AuthorityData = new Abletech.WebApi.Client.Arxivar.Model.AuthorityDataDTO();
            profileDto.Notes = new List<Abletech.WebApi.Client.Arxivar.Model.NoteDTO>();
            profileDto.PaNotes = new List<string>();
            profileDto.PostProfilationActions =
                new List<Abletech.WebApi.Client.Arxivar.Model.PostProfilationActionDTO>();
            profileDto.Document = new Abletech.WebApi.Client.Arxivar.Model.FileDTO() { BufferIds = bufferId };

            if (model.Allegati != null)
            {
                foreach (var allegato in model.Allegati)
                {
                    profileDto.Attachments.AddRange(UploadFile(allegato));
                }
            }

            if (model.AllegatiBin != null)
            {
                foreach (var allegato in model.AllegatiBin)
                {
                    profileDto.Attachments.AddRange(UploadFile(allegato.name, allegato.bytes));
                }
            }

            if (model.allegati_arxivar != null)
            {
                profileDto.Attachments.AddRange(model.allegati_arxivar);
            }

            var classeDoc = profileDto.SetDocumentType(configuration, this._aoo, documentType);

            var status = statesApi.StatesGet(classeDoc.Id);
            profileDto.SetState(model.GetArxivarAttribute().Stato ?? status.First().Id);

            var additional = profileApi.ProfilesGetAdditionalByClasse(
                classeDoc.DocumentType,
                classeDoc.Type2,
                classeDoc.Type3,
                this._aoo
            );
            profileDto.Fields.AddRange(additional);

            profileDto.SetField("DOCNAME", model.DOCNAME);

            if (model.DataDoc.HasValue)
            {
                profileDto.SetField("DataDoc", model.DataDoc.Value);
            }

            if (!string.IsNullOrEmpty(model.User))
            {
                profileDto.SetFromField(GetUserAddressBookEntry(model.User, 1));
            }

            if (!string.IsNullOrEmpty(model.MittenteCodiceRubrica))
            {
                if (model.MittenteIdRubrica.HasValue)
                {
                    profileDto.SetFromField(GetAddressBookEntry(
                        model.MittenteCodiceRubrica,
                        model.MittenteIdRubrica.Value,
                        type: UserProfileType.From
                    ));
                }
            }

            if (model.DestinatariCodiceRubrica != null)
            {
                foreach (var destinatario in model.DestinatariCodiceRubrica)
                {
                    profileDto.SetToField(GetAddressBookEntry(
                        destinatario,
                        model.DestinatariIdRubrica ?? 0,
                        type: UserProfileType.To
                    ));
                }
            }

            model.STATO = model.STATO ?? statesApi.StatesGet(classeDoc.Id).First().Id;
            foreach (var field in model.GetArxivarFields())
            {
                if (field.Value == null) continue;

                if (field.Key.Equals("TO", StringComparison.InvariantCultureIgnoreCase))
                {
                    /*
                    foreach(var to in (List<string>)field.Value)
                    {
                        profileDto.SetToField(GetAddressBookEntry(to, 0, UserProfileType.To));
                    }*/
                }
                else if (field.Key.Equals("FROM", StringComparison.InvariantCultureIgnoreCase))
                {
                    //profileDto.SetFromField(GetAddressBookEntry(field.Value.ToString(), 0, UserProfileType.From));
                }
                else if (field.Key.Equals("CC", StringComparison.InvariantCultureIgnoreCase))
                {
                    //profileDto.SetCcField(GetUserAddressBookEntry(field.Value, 2));
                }
                else if (field.Key.Equals(Attributes.AxFromExternalIdFieldAttribute.AX_KEY))
                {
                }
                else if (field.Key.Equals(Attributes.AxToExternalIdFieldAttribute.AX_KEY))
                {
                }
                else if (field.Key.Equals(Attributes.AxCcExternalIdFieldAttribute.AX_KEY))
                {
                }
                else
                {
                    profileDto.SetField(field.Key, field.Value);
                }
            }


            //profileDto.SetState(model.STATO);

            var newProfile = new Abletech.WebApi.Client.Arxivar.Model.ProfileDTO()
            {
                Fields = profileDto.Fields,
                Document = profileDto.Document,
                Attachments = profileDto.Attachments,
                AuthorityData = profileDto.AuthorityData,
                Notes = profileDto.Notes,
                PaNotes = profileDto.PaNotes,
                PostProfilationActions = profileDto.PostProfilationActions
            };

            if (model.GetArxivarAttribute().Barcode)
            {
                var result = profileApi.ProfilesPostForBarcode(newProfile);
                return result.DocNumber;
            }
            else
            {
                var result = profileApi.ProfilesPost(newProfile);
                return result.DocNumber;
            }
        }

        #endregion

        #region download attachments

        public string[] DownloadAttachments(int docnumber, string outputFolder, bool ignoreException = false)
        {
            Login();
            var attachmentsApi = new Abletech.WebApi.Client.Arxivar.Api.AttachmentsApi(configuration);
            var documentsApi = new Abletech.WebApi.Client.Arxivar.Api.DocumentsApi(configuration);

            var infos = attachmentsApi.AttachmentsGetByDocnumber(docnumber: docnumber);
            var list = new List<string>();
            foreach (var info in infos)
            {
                try
                {
                    var attachment = attachmentsApi.AttachmentsGetById(info.Id);
                    var doc = documentsApi.DocumentsGetForExternalAttachment(info.Id, false);
                    var path = Path.Combine(outputFolder, attachment.Originalname);
                    FileUtils.Write(doc, path).Close();

                    // fix ACL ( permessi ) sul file
                    try
                    {
                        ACUtils.FileUtils.CopyAcl(path);
                    }
                    catch (Exception e)
                    {
                        _logger?.Exception(e);
                    }

                    list.Add(path);
                }
                catch (Exception e)
                {
                    if (!ignoreException) throw;
                    else _logger?.Exception(e);
                }
            }

            return list.ToArray();
        }

        #endregion

        #region download documento

        public (Stream stream, string filename) GetDocumentFileStream(long docnumber, bool forView = false)
        {
            Login();
            var documentsApi = new Abletech.WebApi.Client.Arxivar.Api.DocumentsApi(configuration);

            var response = documentsApi.DocumentsGetForProfileWithHttpInfo(System.Convert.ToInt32(docnumber), forView);

            if (response.Data is FileStream)
            {
                using (response.Data)
                {
                    var fileStream = response.Data as FileStream;
                    var fileName = System.IO.Path.GetFileName(fileStream.Name);
                    MemoryStream memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    fileStream.Close();

                    if (System.IO.File.Exists(fileStream.Name))
                        System.IO.File.Delete(fileStream.Name);
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    return (memoryStream, fileName);
                }
            }
            else
            {
                var fileNameInfo = response.Headers["Content-Disposition"];
                var filename = (new Regex("filename=\"(.*)\"", RegexOptions.IgnoreCase)).Match(fileNameInfo).Groups[0]
                    .Value;
                return (response.Data, filename);
            }
        }

        public string DownloadDocument(long docnumber, string outputFolder, bool forView = false)
        {
            (Stream stream, string filename) = GetDocumentFileStream(docnumber, forView);
            using (stream)
            {
                var fullPath = Path.Combine(outputFolder, filename);
                FileUtils.Write(stream, fullPath).Close();
                return fullPath;
            }
        }

        #endregion

        #region Tasks

        public int Task_ProcessIdFromTaskid(int taskid)
        {
            Login();
            var task = Task_GetByTaskId(taskid);
            return task.ProcessId.Value;
        }

        public IEnumerable<T> Task_GetAttachments<T>(int taskId, string regexpFilterDoctype = null)
            where T : AXModel<T>, new()
        {
            Login();

            var processId = Task_ProcessIdFromTaskid(taskId);

            var profileApi = new Abletech.WebApi.Client.Arxivar.Api.ProfilesApi(configuration);
            var taskWorkAttachmentsV2Api =
                new Abletech.WebApi.Client.Arxivar.Api.TaskWorkAttachmentsV2Api(configuration);
            var attachments =
                taskWorkAttachmentsV2Api.TaskWorkAttachmentsV2GetAttachmentsByProcessId(processId) as JObject;
            var targetDocType = String.IsNullOrEmpty(regexpFilterDoctype)
                ? Regex.Escape((new T()).GetArxivarAttribute().DocumentType)
                : regexpFilterDoctype;
            Regex targetDocRx = new Regex(targetDocType, RegexOptions.IgnoreCase);

            int i = 0;
            int docnumber_pos = -1;
            int tipoallegato_pos = -1;

            foreach (var c in attachments["columns"])
            {
                var id = (string)c["id"];
                if (id == "DOCNUMBER")
                {
                    docnumber_pos = i;
                }

                if (id == "TIPOALLEGATO")
                {
                    tipoallegato_pos = i;
                }

                i++;
            }

            foreach (var row in attachments["data"])
            {
                var tipo_allegato = row[tipoallegato_pos].Value<int?>();
                var v = row[docnumber_pos];
                if (tipo_allegato != 2) continue;
                if (!v.Value<int?>().HasValue) continue;
                var docnumber = v.Value<int>();
                try
                {
                    var profile = profileApi.ProfilesGetSchema(docnumber, false);
                    var cDocType = profile.GetDocumentType();
                    if (!targetDocRx.Match(cDocType).Success) continue;
                }
                catch
                {
                    // se il documento non è accessibile continua
                    // TODO: migliorare questa logica 
                    continue;
                }

                yield return this.GetProfile<T>(docnumber);
            }
        }

        public IEnumerable<(int docnumber, T documento)> Task_GetDocument<T>(int taskId) where T : AXModel<T>, new()
        {
            var processId = Task_ProcessIdFromTaskid(taskId);

            // TaskWorkV2_getDocumentsByProcessId
            Login();
            var profileApi = new Abletech.WebApi.Client.Arxivar.Api.ProfilesApi(configuration);
            var taskWorkV2Api = new Abletech.WebApi.Client.Arxivar.Api.TaskWorkV2Api(configuration);

            var targetDocType = (new T()).GetArxivarAttribute().DocumentType;

            var select = taskWorkV2Api.TaskWorkV2GetDefaultSelect();
            select.Fields.Select("DOCNUMBER");

            var docs = taskWorkV2Api.TaskWorkV2GetDocumentsByProcessId(processId, select) as JObject;
            var docnumber_pos = -1;
            var s = docs["columns"].AsEnumerable().Select((d, y) => (d, y));
            int i = 0;
            foreach (var c in docs["columns"])
            {
                if ((string)c["id"] == "DOCNUMBER")
                {
                    docnumber_pos = i;
                    break;
                }

                i++;
            }

            foreach (var row in docs["data"])
            {
                var docnumber = (int)row[docnumber_pos];
                var profile = profileApi.ProfilesGetSchema(docnumber, false);
                var cDocType = profile.GetDocumentType();
                if (targetDocType != cDocType) continue;
                yield return (docnumber, this.GetProfile<T>(docnumber));
            }
        }

        public Abletech.WebApi.Client.Arxivar.Model.TaskWorkDTO Task_GetByTaskId(int taskid)
        {
            Login();
            var taskWorkV2Api = new Abletech.WebApi.Client.Arxivar.Api.TaskWorkV2Api(configuration);
            return taskWorkV2Api.TaskWorkV2GetTaskWorkById(taskid);
        }

        public void Task_AggiungiAllegato(int taskWorkId, string filePath, string filename = null)
        {
            Login();
            var taskWorkAttachV2Api = new Abletech.WebApi.Client.Arxivar.Api.TaskWorkAttachmentsV2Api(configuration);
            var bufferId = UploadFile(filePath, filename: filename).First();
            taskWorkAttachV2Api.TaskWorkAttachmentsV2AddNewExternalAttachments(bufferId: bufferId,
                taskWorkId: taskWorkId);
        }

        public long Task_GetUserIdOfTaskId(int processId, int taskWorkId)
        {
            Login();
            var taskHistoryApi = new Abletech.WebApi.Client.Arxivar.Api.TaskWorkHistoryApi(configuration);
            var taskHistory = taskHistoryApi.TaskWorkHistoryGetHistoryByProcessId(processId);
            var userId = (from task in taskHistory
                where task.GetValue<long>("ID") == taskWorkId
                select task.Columns.GetValue<long>("UTENTE")).First();
            return userId;
        }

        #endregion

        #region fascioli

        public List<Abletech.WebApi.Client.Arxivar.Model.RowSearchResult> GetFascicoloDocuments(int id)
        {
            Login();

            var foldersApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);
            var searchApi = new Abletech.WebApi.Client.Arxivar.Api.SearchesApi(configuration);
            var select = searchApi.SearchesGetSelect();
            select.Fields.Select("CLASSEDOC");
            return foldersApi.FoldersGetDocumentsById(id, select);
        }

        public int GetFascicoloLevel(int id)
        {
            Login();
            var folderApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);
            var folderInfo = folderApi.FoldersGetById(id);
            return folderInfo.FullPath.Count(f => f == '\\');
        }

        public List<Abletech.WebApi.Client.Arxivar.Model.FolderDTO> GetFascoloFiglio(int id, string name)
        {
            Login();
            var foldersApi2 = new Abletech.WebApi.Client.Arxivar.Api.FoldersV2Api(configuration);
            return foldersApi2.FoldersV2FindInFolderByName(id, name);
            var foldersApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);
            return foldersApi.FoldersFindInFolderByName(id, name);
        }


        public List<int> FascicoliGetByDocnumber(int docnumber)
        {
            Login();
            var foldersApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);
            var folders = foldersApi.FoldersFindByDocnumber(docnumber);
            return folders.Where(f => f.Id.HasValue).Select(f => f.Id.Value).ToList();
        }

        public int FascicoliFolderExists(int parentFolder, string subfolderName)
        {
            Login();
            var foldersApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);

            var folders = foldersApi.FoldersGetByParentId(parentFolder);

            var folderSearch = folders.Where(f => f.Name.ToLower().Equals(subfolderName.ToLower()));
            if (folderSearch.Any())
            {
                return folderSearch.First().Id.GetValueOrDefault();
            }
            else
            {
                return 0;
            }
        }

        public int FascicoliCreateFolder(int parentFolder, string subfodlerName)
        {
            Login();
            var foldersApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);
            int subfodler = FascicoliFolderExists(parentFolder, subfodlerName);
            if (subfodler == 0) // se non esiste
            {
                var newfodler = foldersApi.FoldersNew(parentFolder, subfodlerName);
                subfodler = newfodler.Id.Value;
            }

            return subfodler;
        }

        public void FascicoliMoveToFolder(int folderId, int docnumber)
        {
            Login();
            var foldersApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);
            foldersApi.FoldersInsertDocnumbers(folderId, new List<int?>() { docnumber });
        }

        public int FascicoliMoveToSubfolder(int parentFolder, string subfolderName, int docnumber)
        {
            Login();
            var foldersApi = new Abletech.WebApi.Client.Arxivar.Api.FoldersApi(configuration);

            var folders = foldersApi.FoldersGetByParentId(parentFolder);

            var folderSearch = folders.Where(f => f.Name.ToLower().Equals(subfolderName.ToLower()));
            var folder_exists = folderSearch.Any();

            var folderId = FascicoliCreateFolder(parentFolder, subfolderName);

            // rimuovi dalla cartella precedente
            try
            {
                foldersApi.FoldersRemoveDocumentsInFolder(parentFolder, new List<int?>() { docnumber });
            }
            catch
            {
            }

            // rimuove se già presente 
            try
            {
                foldersApi.FoldersRemoveDocumentsInFolder(folderId, new List<int?>() { docnumber });
            }
            catch
            {
            }

            // aggiungi alla cartella di destinazione
            FascicoliMoveToFolder(folderId, docnumber);
            return folderId;
        }

        #endregion

        #region users

        public Abletech.WebApi.Client.Arxivar.Model.UserProfileDTO GetUserAddressBookEntry(string username,
            int type = 0)
        {
            Login();

            var addressBookApi = new Abletech.WebApi.Client.Arxivar.Api.AddressBookApi(configuration);
            var userApi = new Abletech.WebApi.Client.Arxivar.Api.UsersApi(configuration);
            // Dm_Rubrica.RAGIONE_SOCIALE = username | Dm_Rubrica.TIPO=U | Dm_Rubrica.STATO=P
            var users = userApi.UsersGet_0();
            var search = users.Where(u =>
                u.Description.Equals(username, StringComparison.CurrentCultureIgnoreCase) ||
                u.CompleteName.Equals(username, StringComparison.CurrentCultureIgnoreCase)
            );
            var id = search.FirstOrDefault()?.User;
            return addressBookApi.AddressBookGetByUserId(id, type);
        }


        public Abletech.WebApi.Client.Arxivar.Model.UserInfoDTO UserInfo()
        {
            Login();
            var api = new Abletech.WebApi.Client.Arxivar.Api.UsersApi(configuration);
            var userInfo = api.UsersGetUserInfo();
            return userInfo;
        }


        public List<Abletech.WebApi.Client.Arxivar.Model.UserCompleteDTO> Users()
        {
            Login();
            var userApi = new Abletech.WebApi.Client.Arxivar.Api.UsersApi(configuration);
            return userApi.UsersGet_0();
        }


        public Abletech.WebApi.Client.Arxivar.Model.UserInfoDTO UserGet(string aoo, string username)
        {
            Login();
            var userSearchApi = new Abletech.WebApi.Client.Arxivar.Api.UserSearchApi(configuration);

            var select = userSearchApi.UserSearchGetSelect()
                .Select("UTENTE");

            var search = userSearchApi.UserSearchGetSearch()
                .SetString("DESCRIPTION", username)
                .SetString("AOO", aoo);

            var result = userSearchApi
                .UserSearchPostSearch(
                    new Abletech.WebApi.Client.Arxivar.Model.UserSearchCriteriaDTO(selectDto: select,
                        searchDto: search)).FirstOrDefault();
            if (result == null)
            {
                throw new NotFoundException($"user '{aoo}\\{username}' not found");
            }

            var userApi = new Abletech.WebApi.Client.Arxivar.Api.UsersApi(configuration);
            var userId = result.GetValue<int>("UTENTE");
            return userApi.UsersGet(userId);
        }

        public bool UserExists(string aoo, string username)
        {
            try
            {
                UserGet(aoo, username);
                return true;
            }
            catch (NotFoundException)
            {
                return false;
            }
        }

        public bool UserCreate(
            string username,
            string aoo,
            string description,
            string defaultPassword,
            string email = null,
            string lang = "IT",
            int tipo = 1,
            bool mustChangePassword = true,
            bool workflow = true,
            IEnumerable<string> groups = null
        )
        {
            this._logger?.Information($"Creazione utente {username}");
            Login();
            LoginManagment();
            var userApi = new Abletech.WebApi.Client.Arxivar.Api.UsersApi(configuration);
            var usersManagementApi =
                new Abletech.WebApi.Client.ArxivarManagement.Api.UsersManagementApi(configurationManagement);

            var newUser = userApi.UsersInsert(
                new Abletech.WebApi.Client.Arxivar.Model.UserInsertDTO()
                {
                    Password = defaultPassword,
                    Description = username,
                    CompleteDescription = description,
                    Email = email,
                    Workflow = workflow,
                    MustChangePassword = mustChangePassword ? 1 : 0,
                    PasswordChange = true,
                    Type = tipo,
                    Viewer = 0,
                    Group = 2,
                    UserState = 1,
                    BusinessUnit = aoo,
                    Lang = lang,
                    DefaultType = 0,
                    Type2 = 0,
                    Type3 = 0,
                }
            );


            if (groups != null)
            {
                var existingGroups = userApi.UsersGetGroups();
                var newGroups = existingGroups.Where(group =>
                        groups.Select(g => g.ToLower()).Contains(group.CompleteName.ToLower()) ||
                        groups.Select(g => g.ToLower()).Contains(group.Description.ToLower())
                    )
                    .Select(g =>
                        new Abletech.WebApi.Client.ArxivarManagement.Model.UserSimpleDTO(user: g.Id,
                            description: g.Description))
                    .ToList();
                usersManagementApi.UsersManagementSetUserGroups(userId: newUser.User, groups: newGroups);
            }

            return true;
        }

        public bool UserCreateIfNotExists(
            string username,
            string aoo,
            string description,
            string defaultPassword,
            string email = null,
            string lang = "it",
            int tipo = 1,
            bool mustChangePassword = true,
            bool workflow = true,
            IEnumerable<string> groups = null
        )
        {
            if (!UserExists(aoo, username))
            {
                return UserCreate(
                    username: username,
                    aoo: aoo,
                    description: description,
                    defaultPassword: defaultPassword,
                    email: email,
                    lang: lang,
                    tipo: tipo,
                    mustChangePassword: mustChangePassword,
                    workflow: workflow,
                    groups: groups
                );
            }

            foreach (var group in groups)
            {
                UserAddGroup(aoo, username, group);
            }

            return false;
        }

        public bool UserAddGroup(string aoo, string username, string groupName)
        {
            Login();
            LoginManagment();
            var usersManagementApi =
                new Abletech.WebApi.Client.ArxivarManagement.Api.UsersManagementApi(configurationManagement);
            var userApi = new Abletech.WebApi.Client.Arxivar.Api.UsersApi(configuration);

            var user = UserGet(aoo, username);
            var existingGroups = userApi.UsersGetGroups();
            var group = existingGroups.FirstOrDefault(g =>
                g.Description.Equals(groupName, StringComparison.CurrentCultureIgnoreCase)
            );
            if (group == null)
            {
                throw new NotFoundException($"Arxivar group '{groupName}' not found");
            }

            var userGroups = usersManagementApi.UsersManagementGetUserGroups(user.User);
            if (!userGroups.Any(g => g.Description.Equals(groupName, StringComparison.CurrentCultureIgnoreCase)))
            {
                userGroups.Add(
                    new Abletech.WebApi.Client.ArxivarManagement.Model.UserSimpleDTO(user: group.Id,
                        description: group.Description));
                usersManagementApi.UsersManagementSetUserGroups(userId: user.User, groups: userGroups);
                return true;
            }

            return false;
        }

        public void UserUpdate(Abletech.WebApi.Client.Arxivar.Model.UserCompleteDTO user)
        {
            Login();
            var userApi = new Abletech.WebApi.Client.Arxivar.Api.UsersApi(configuration);
            var update = new Abletech.WebApi.Client.Arxivar.Model.UserUpdateDTO()
            {
                User = user.User,
                Group = user.Group,
                Description = user.Description,
                Email = user.Email,
                BusinessUnit = user.BusinessUnit,
                Password = user.Password,
                DefaultType = user.DefaultType,
                Type2 = user.Type2,
                Type3 = user.Type3,
                InternalFax = user.InternalFax,
                LastMail = user.LastMail,
                Category = user.Category,
                Workflow = user.Workflow,
                DefaultState = user.DefaultState,
                AddressBook = user.AddressBook,
                UserState = user.UserState,
                MailServer = user.MailServer,
                WebAccess = user.WebAccess,
                Upload = user.Upload,
                Folders = user.Folders,
                Flow = user.Flow,
                Sign = user.Sign,
                Viewer = user.Viewer,
                Protocol = user.Protocol,
                Models = user.Models,
                Domain = user.Domain,
                OutState = user.OutState,
                MailBody = user.MailBody,
                Notify = user.Notify,
                MailClient = user.MailClient,
                HtmlBody = user.HtmlBody,
                RespAos = user.RespAos,
                AssAos = user.AssAos,
                CodFis = user.CodFis,
                Pin = user.Pin,
                Guest = user.Guest,
                PasswordChange = !string.IsNullOrEmpty(user.Password),
                Marking = user.Marking,
                Type = user.Type,
                MailOutDefault = user.MailOutDefault,
                BarcodeAccess = user.BarcodeAccess,
                MustChangePassword = user.MustChangePassword,
                Lang = user.Lang,
                Ws = user.Ws,
                DisablePswExpired = user.DisablePswExpired,
                CompleteDescription = user.CompleteDescription,
                CanAddVirtualStamps = user.CanAddVirtualStamps,
                CanApplyStaps = user.CanApplyStaps,
            };
            userApi.UsersUpdate(user.User, update);
        }

        #endregion

        #region workflow

        public void DeleteWorkflow(int? processId)
        {
            var workflowApi = new Abletech.WebApi.Client.Arxivar.Api.WorkflowApi(configuration);
            workflowApi.WorkflowStopWorkflow(processId.Value);
            workflowApi.WorkflowDeleteWorkflow(processId, true);
            workflowApi.WorkflowFreeUserConstraint(processId.Value);
        }

        #endregion

        #region workflow2

        public void Wf2_SetVariable(Guid taskId, string name, string value)
        {
            var taskOperationsApi =
                new Abletech.WebApi.Client.ArxivarWorkflow.Api.TaskOperationsApi(configurationWorkflow);

            var response = taskOperationsApi.ApiV1TaskOperationsTaskTaskIdVariablesGet(taskId);

            var variable = response.First(v => v.VariableDefinition.Configuration.Name.Equals(name));

            taskOperationsApi.ApiV1TaskOperationsExecuteSetVariablesPost(
                new Abletech.WebApi.Client.ArxivarWorkflow.Model.ExecuteSetVariablesOperationRm(
                    setVariables: new List<Abletech.WebApi.Client.ArxivarWorkflow.Model.ProcessManualSetVariableRm>()
                    {
                        {
                            new Abletech.WebApi.Client.ArxivarWorkflow.Model.ProcessManualSetVariableRm()
                        }
                    }
                )
                {
                });
        }

        #endregion
    }
}