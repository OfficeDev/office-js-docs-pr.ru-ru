---
title: Элемент WebApplicationInfo в файле манифеста
description: Справочная документация по элементу WebApplicationInfo для файлов манифеста надстроек Office (XML).
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 8644529d82204cb9fbc07c6fe9f8a35b60a512c8
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293809"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="d29d3-103">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="d29d3-103">WebApplicationInfo element</span></span>

<span data-ttu-id="d29d3-104">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="d29d3-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="d29d3-105">*Ресурс* OAuth 2,0, которому клиентским приложениям Office могут потребоваться разрешения.</span><span class="sxs-lookup"><span data-stu-id="d29d3-105">An OAuth 2.0 *resource* to which the Office client application might need permissions.</span></span>
- <span data-ttu-id="d29d3-106">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="d29d3-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="d29d3-107">В настоящее время API единого входа поддерживается для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d29d3-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="d29d3-108">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d29d3-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="d29d3-109">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="d29d3-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="d29d3-110">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="d29d3-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="d29d3-111">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="d29d3-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="d29d3-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d29d3-112">Child elements</span></span>

|  <span data-ttu-id="d29d3-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="d29d3-113">Element</span></span> |  <span data-ttu-id="d29d3-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d29d3-114">Required</span></span>  |  <span data-ttu-id="d29d3-115">Описание</span><span class="sxs-lookup"><span data-stu-id="d29d3-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d29d3-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="d29d3-116">**Id**</span></span>    |  <span data-ttu-id="d29d3-117">Да</span><span class="sxs-lookup"><span data-stu-id="d29d3-117">Yes</span></span>   |  <span data-ttu-id="d29d3-118">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="d29d3-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="d29d3-119">**мсаид**</span><span class="sxs-lookup"><span data-stu-id="d29d3-119">**MsaId**</span></span>    |  <span data-ttu-id="d29d3-120">Нет</span><span class="sxs-lookup"><span data-stu-id="d29d3-120">No</span></span>   |  <span data-ttu-id="d29d3-121">Идентификатор клиента веб-приложения надстройки для MSA, зарегистрированного в msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="d29d3-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="d29d3-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="d29d3-122">**Resource**</span></span>  |  <span data-ttu-id="d29d3-123">Да</span><span class="sxs-lookup"><span data-stu-id="d29d3-123">Yes</span></span>   |  <span data-ttu-id="d29d3-124">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="d29d3-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="d29d3-125">Scopes</span><span class="sxs-lookup"><span data-stu-id="d29d3-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="d29d3-126">Да</span><span class="sxs-lookup"><span data-stu-id="d29d3-126">Yes</span></span>  |  <span data-ttu-id="d29d3-127">Задает разрешения, необходимые надстройке для ресурса, например Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="d29d3-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="d29d3-128">Authorizations</span><span class="sxs-lookup"><span data-stu-id="d29d3-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="d29d3-129">Нет</span><span class="sxs-lookup"><span data-stu-id="d29d3-129">No</span></span>   | <span data-ttu-id="d29d3-130">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="d29d3-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="d29d3-131">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="d29d3-131">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
