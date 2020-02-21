---
title: Получение вложений в надстройке Outlook
description: Надстройка может использовать API вложений для отправки информации о вложениях удаленной службе.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 7188359193e675f53d0e8358c75f03669b34a170
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166767"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a>Получение вложений элемента Outlook с сервера

Надстройка Outlook не может передавать вложения для выбранного элемента непосредственно в удаленную службу, работающую на сервере. Вместо этого она может использовать API вложений для отправки информации о вложениях в такую удаленную службу. Затем эта служба может обратиться напрямую к серверу Exchange для получения вложений.

Чтобы отправить информацию о вложениях в удаленную службу, используйте следующие свойства и функцию:

- Свойство [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities): предоставляет URL-адрес веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик. Служба использует этот URL-адрес, чтобы вызвать метод [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) или операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) для EWS.

- Свойство [Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): получает массив объектов [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) (по одному для каждого вложения в элемент).

- Функция [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods): асинхронно вызывает сервер Exchange Server с почтовым ящиком, чтобы получить маркер обратного вызова, который клиентский сервер отправит обратно на сервер Exchange Server для проверки подлинности запроса на получение вложения.

## <a name="using-the-attachments-api"></a>Использование API вложений

Чтобы использовать API вложений для получения вложений из почтового ящика Exchange, выполните следующие действия.

1. Отобразите надстройку, когда пользователь просматривает сведения о встрече или сообщение, которые содержат вложение.

1. Получите маркер обратного вызова с сервера Exchange.

1. Отправьте маркер обратного вызова и сведения о вложениях в удаленную службу.

1. Получите вложения с сервера Exchange с помощью метода `ExchangeService.GetAttachments` или операции `GetAttachment`.

Каждый из этих шагов рассматривается в следующих разделах более подробно на примере кода [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments).

> [!NOTE]
> Код в этих примерах был сокращен, чтобы уделить основное внимание информации о вложениях. Пример содержит дополнительный код для проверки подлинности надстройки на удаленном сервере и управления состоянием запроса.

## <a name="get-a-callback-token"></a>Получение маркера обратного вызова

Объект [Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) предоставляет функцию `getCallbackTokenAsync` для получения маркера, с помощью которого удаленный сервер может пройти проверку подлинности на сервере Exchange Server. В приведенном ниже фрагменте кода показаны функция надстройки, отправляющая асинхронный запрос маркера обратного вызова, и функция обратного вызова, получающая ответ. Маркер обратного вызова хранится в объекте запроса к службе, определяемом в следующем разделе.

```js
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "") {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
}

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}
```

## <a name="send-attachment-information-to-the-remote-service"></a>Отправка сведений о вложениях в удаленную службу

От удаленной службы, которую вызывает ваша надстройка, зависит способ отправки информации о вложениях в эту службу. В данном примере такой удаленной службой является приложение веб-API, созданное с помощью Visual Studio 2013. Удаленная служба ожидает получения сведений о вложениях в объекте JSON. Следующий код инициализирует объект, содержащий информацию о вложениях.

```js
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
 var serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };
```

<br/>

Свойство `Office.context.mailbox.item.attachments` содержит коллекцию объектов `AttachmentDetails`, по одному на каждое вложение в элементе. В большинстве случаев надстройке достаточно передать в удаленную службу свойство объекта `AttachmentDetails`, содержащее идентификатор вложения. Если удаленной службе нужно больше сведений о вложении, вы можете полностью или частично передать объект `AttachmentDetails`. Приведенный ниже код определяет метод, который помещает весь массив `AttachmentDetails` в объект `serviceRequest` и отправляет запрос в удаленную службу.

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (var i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      var names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (i = 0; i < response.attachmentNames.length; i++) {
        names += response.attachmentNames[i] + "<br />";
      }
      document.getElementById("names").innerHTML = names;
    } else {
      app.showNotification("Runtime error", response.message);
    }
  }).fail(function (status) {

  }).always(function () {
    $('.disable-while-sending').prop('disabled', false);
  })
}
```

## <a name="get-the-attachments-from-the-exchange-server"></a>Получение вложений с сервера Exchange Server

Ваша удаленная служба может использовать метод [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) управляемого API веб-служб Exchange или операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) управляемого API веб-служб Exchange для получения вложений с сервера. Приложению-службе необходимы два объекта, чтобы выполнить десериализацию строки JSON в объекты .NET Framework, которые можно использовать на сервере. В следующем коде показаны определения объектов десериализации.

```cs
namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a>Использование управляемого API EWS для получения вложений

Если вы используете в своей удаленной службе [управляемый API EWS](https://go.microsoft.com/fwlink/?LinkID=255472), вы можете воспользоваться методом [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange), который создаст, отправит и получит SOAP-запрос EWS для получения вложений. Рекомендуем использовать управляемый API EWS, поскольку он требует меньше строк кода и обеспечивает более интуитивный интерфейс для вызовов EWS. Приведенный ниже код отправляет один запрос на получение всех вложений, а также возвращает количество и имена обработанных вложений.

```cs
private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  // Create an ExchangeService object, set the credentials and the EWS URL.
  ExchangeService service = new ExchangeService();
  service.Credentials = new OAuthCredentials(request.attachmentToken);
  service.Url = new Uri(request.ewsUrl);

  var attachmentIds = new List<string>();

  foreach (AttachmentDetails attachment in request.attachments)
  {
    attachmentIds.Add(attachment.id);
  }

  // Call the GetAttachments method to retrieve the attachments on the message.
  // This method results in a GetAttachments EWS SOAP request and response
  // from the Exchange server.
  var getAttachmentsResponse =
    service.GetAttachments(attachmentIds.ToArray(),
                            null,
                            new PropertySet(BasePropertySet.FirstClassProperties,
                                            ItemSchema.MimeContent));

  if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
  {
    foreach (var attachmentResponse in getAttachmentsResponse)
    {
      attachmentNames.Add(attachmentResponse.Attachment.Name);

      // Write the content of each attachment to a stream.
      if (attachmentResponse.Attachment is FileAttachment)
      {
        FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
        Stream s = new MemoryStream(fileAttachment.Content);
        // Process the contents of the attachment here.
      }

      if (attachmentResponse.Attachment is ItemAttachment)
      {
        ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
        Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
        // Process the contents of the attachment here.
      }

      attachmentsProcessedCount++;
    }
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

### <a name="use-ews-to-get-the-attachments"></a>Использование EWS для получения вложений

Если в удаленной службе используется EWS, для получения вложений с сервера Exchange Server необходимо создать SOAP-запрос [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation). Приведенный ниже код возвращает строку, представляющую SOAP-запрос. Удаленная служба вставляет в строку идентификатор вложения, используя метод `String.Format`.


```cs
private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""https://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""https://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
```

<br/>

Приведенный ниже метод получает вложения с сервера Exchange Server, использует запрос EWS `GetAttachment`. При такой реализации отправляется отдельный запрос для каждого вложения и возвращается количество обработанных вложений. Каждый ответ обрабатывается в отдельном методе `ProcessXmlResponse`, определение которого представлено ниже.

```cs
private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  foreach (var attachment in request.attachments)
  {
    // Prepare a web request object.
    HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
    webRequest.Headers.Add("Authorization",
      string.Format("Bearer {0}", request.attachmentToken));
    webRequest.PreAuthenticate = true;
    webRequest.AllowAutoRedirect = false;
    webRequest.Method = "POST";
    webRequest.ContentType = "text/xml; charset=utf-8";

    // Construct the SOAP message for the GetAttachment operation.
    byte[] bodyBytes = Encoding.UTF8.GetBytes(
      string.Format(GetAttachmentSoapRequest, attachment.id));
    webRequest.ContentLength = bodyBytes.Length;

    Stream requestStream = webRequest.GetRequestStream();
    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
    requestStream.Close();

    // Make the request to the Exchange server and get the response.
    HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

    // If the response is okay, create an XML document from the response
    // and process the request.
    if (webResponse.StatusCode == HttpStatusCode.OK)
    {
      var responseStream = webResponse.GetResponseStream();

      var responseEnvelope = XElement.Load(responseStream);

      // After creating a memory stream containing the contents of the
      // attachment, this method writes the XML document to the trace output.
      // Your service would perform it's processing here.
      if (responseEnvelope != null)
      {
        var processResult = ProcessXmlResponse(responseEnvelope);
        attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

      }

      // Close the response stream.
      responseStream.Close();
      webResponse.Close();

    }
    // If the response is not OK, return an error message for the
    // attachment.
    else
    {
      var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
        "Error message: {1}.", attachment.name, webResponse.StatusDescription);
      attachmentNames.Add(errorString);
    }
    attachmentsProcessedCount++;
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

<br/>

Каждый ответ на операцию `GetAttachment` отправляется в метод `ProcessXmlResponse`. Этот метод проверяет ответ на наличие ошибок. Если ошибки не найдены, он обрабатывает вложенные файлы и элементы. Метод `ProcessXmlResponse` выполняет большую часть операций по обработке вложения.

```cs
// This method processes the response from the Exchange server.
// In your application the bulk of the processing occurs here.
private string ProcessXmlResponse(XElement responseEnvelope)
{
  // First, check the response for web service errors.
  var errorCodes = from errorCode in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                    select errorCode;
  // Return the first error code found.
  foreach (var errorCode in errorCodes)
  {
    if (errorCode.Value != "NoError")
    {
      return string.Format("Could not process result. Error: {0}", errorCode.Value);
    }
  }

  // No errors found, proceed with processing the content.
  // First, get and process file attachments.
  var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                        select fileAttachment;
  foreach(var fileAttachment in fileAttachments)
  {
    var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
    var fileData = System.Convert.FromBase64String(fileContent.Value);
    var s = new MemoryStream(fileData);
    // Process the file attachment here.
  }

  // Second, get and process item attachments.
  var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                        select itemAttachment;
  foreach(var itemAttachment in itemAttachments)
  {
    var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
    if (message != null)
    {
      // Process a message here.
      break;
    }
    var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
    if (calendarItem != null)
    {
      // Process calendar item here.
      break;
    }
    var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
    if (contact != null)
    {
      // Process contact here.
      break;
    }
    var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
    if (task != null)
    {
      // Process task here.
      break;
    }
    var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
    if (meetingMessage != null)
    {
      // Process meeting message here.
      break;
    }
    var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
    if (meetingRequest != null)
    {
      // Process meeting request here.
      break;
    }
    var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
    if (meetingResponse != null)
    {
      // Process meeting response here.
      break;
    }
    var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
    if (meetingCancellation != null)
    {
      // Process meeting cancellation here.
      break;
    }
  }

  return string.Empty;
}
```

## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм чтения](read-scenario.md)
- [Сведения об управляемом API EWS, EWS и веб-службах в Exchange](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [Начало работы с клиентскими приложениями, использующими управляемый API EWS](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [Надстройка Outlook для примера AttachmentsDemo](https://github.com/OfficeDev/outlook-add-in-attachments-demo)
