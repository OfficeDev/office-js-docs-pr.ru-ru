# <a name="webapplicationinfo-element"></a><span data-ttu-id="57561-101">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="57561-101">WebApplicationInfo element</span></span>

<span data-ttu-id="57561-102">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="57561-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="57561-103">*Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="57561-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="57561-104">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="57561-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="57561-105">В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="57561-105">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="57561-106">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы требований API Identity](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="57561-106">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span> <span data-ttu-id="57561-107">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для области клиентов Office 365.</span><span class="sxs-lookup"><span data-stu-id="57561-107">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="57561-108">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="57561-108">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="57561-109">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="57561-109">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="57561-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="57561-110">Child elements</span></span>

|  <span data-ttu-id="57561-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="57561-111">Element</span></span> |  <span data-ttu-id="57561-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="57561-112">Required</span></span>  |  <span data-ttu-id="57561-113">Описание</span><span class="sxs-lookup"><span data-stu-id="57561-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="57561-114">**Id**</span><span class="sxs-lookup"><span data-stu-id="57561-114">**Id**</span></span>    |  <span data-ttu-id="57561-115">Да</span><span class="sxs-lookup"><span data-stu-id="57561-115">Yes</span></span>   |  <span data-ttu-id="57561-116">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="57561-116">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="57561-117">**Resource**</span><span class="sxs-lookup"><span data-stu-id="57561-117">**Resource**</span></span>  |  <span data-ttu-id="57561-118">Да</span><span class="sxs-lookup"><span data-stu-id="57561-118">Yes</span></span>   |  <span data-ttu-id="57561-119">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="57561-119">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="57561-120">Scopes</span><span class="sxs-lookup"><span data-stu-id="57561-120">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="57561-121">Нет</span><span class="sxs-lookup"><span data-stu-id="57561-121">No</span></span>  |  <span data-ttu-id="57561-122">Указывает разрешения, необходимые надстройке для работы с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="57561-122">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="57561-123">В настоящее время необходимо, чтобы ресурс надстройки соответствовал ее ведущему приложению.</span><span class="sxs-lookup"><span data-stu-id="57561-123">Note: Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="57561-124">Office запрашивает маркер для надстройки, только если может подтвердить право собственности. В настоящее время для этого необходимо, чтобы надстройка размещалась под полным доменным именем ресурса.</span><span class="sxs-lookup"><span data-stu-id="57561-124">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="57561-125">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="57561-125">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
