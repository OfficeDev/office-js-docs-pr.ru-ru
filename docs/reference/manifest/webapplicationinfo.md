---
title: Элемент WebApplicationInfo в файле манифеста
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: e10aee1bf3fb99099d282acd428fa0348229701c
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477868"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="3f0cb-102">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="3f0cb-102">WebApplicationInfo element</span></span>

<span data-ttu-id="3f0cb-103">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="3f0cb-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="3f0cb-104">*Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="3f0cb-105">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="3f0cb-106">В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="3f0cb-107">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="3f0cb-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="3f0cb-108">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="3f0cb-109">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="3f0cb-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="3f0cb-110">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="3f0cb-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="3f0cb-111">Child elements</span></span>

|  <span data-ttu-id="3f0cb-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="3f0cb-112">Element</span></span> |  <span data-ttu-id="3f0cb-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="3f0cb-113">Required</span></span>  |  <span data-ttu-id="3f0cb-114">Описание</span><span class="sxs-lookup"><span data-stu-id="3f0cb-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3f0cb-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="3f0cb-115">**Id**</span></span>    |  <span data-ttu-id="3f0cb-116">Да</span><span class="sxs-lookup"><span data-stu-id="3f0cb-116">Yes</span></span>   |  <span data-ttu-id="3f0cb-117">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="3f0cb-118">**мсаид**</span><span class="sxs-lookup"><span data-stu-id="3f0cb-118">**MsaId**</span></span>    |  <span data-ttu-id="3f0cb-119">Нет</span><span class="sxs-lookup"><span data-stu-id="3f0cb-119">No</span></span>   |  <span data-ttu-id="3f0cb-120">Идентификатор клиента веб-приложения надстройки для MSA, зарегистрированного в msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-120">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="3f0cb-121">**Resource**</span><span class="sxs-lookup"><span data-stu-id="3f0cb-121">**Resource**</span></span>  |  <span data-ttu-id="3f0cb-122">Да</span><span class="sxs-lookup"><span data-stu-id="3f0cb-122">Yes</span></span>   |  <span data-ttu-id="3f0cb-123">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-123">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="3f0cb-124">Scopes</span><span class="sxs-lookup"><span data-stu-id="3f0cb-124">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="3f0cb-125">Да</span><span class="sxs-lookup"><span data-stu-id="3f0cb-125">Yes</span></span>  |  <span data-ttu-id="3f0cb-126">Задает разрешения, необходимые надстройке для ресурса, например Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-126">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="3f0cb-127">Авторизации</span><span class="sxs-lookup"><span data-stu-id="3f0cb-127">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="3f0cb-128">Нет</span><span class="sxs-lookup"><span data-stu-id="3f0cb-128">No</span></span>   | <span data-ttu-id="3f0cb-129">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="3f0cb-129">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="3f0cb-130">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="3f0cb-130">WebApplicationInfo example</span></span>

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
