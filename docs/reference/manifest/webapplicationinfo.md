---
title: Элемент WebApplicationInfo в файле манифеста
description: Справочная документация по элементу VersionOverrides для файлов манифеста надстроек Office (XML).
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 5be75c6e202e40d60961a1b930ef43e583dee240
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094409"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="dc753-103">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="dc753-103">WebApplicationInfo element</span></span>

<span data-ttu-id="dc753-104">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="dc753-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="dc753-105">*Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="dc753-105">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="dc753-106">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="dc753-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="dc753-107">В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="dc753-107">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="dc753-108">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="dc753-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="dc753-109">Если вы работаете с надстройкой Outlook, обязательно включите современная проверка подлинности для клиента Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="dc753-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="dc753-110">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="dc753-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="dc753-111">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="dc753-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="dc753-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="dc753-112">Child elements</span></span>

|  <span data-ttu-id="dc753-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="dc753-113">Element</span></span> |  <span data-ttu-id="dc753-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="dc753-114">Required</span></span>  |  <span data-ttu-id="dc753-115">Описание</span><span class="sxs-lookup"><span data-stu-id="dc753-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="dc753-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="dc753-116">**Id**</span></span>    |  <span data-ttu-id="dc753-117">Да</span><span class="sxs-lookup"><span data-stu-id="dc753-117">Yes</span></span>   |  <span data-ttu-id="dc753-118">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="dc753-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="dc753-119">**мсаид**</span><span class="sxs-lookup"><span data-stu-id="dc753-119">**MsaId**</span></span>    |  <span data-ttu-id="dc753-120">Нет</span><span class="sxs-lookup"><span data-stu-id="dc753-120">No</span></span>   |  <span data-ttu-id="dc753-121">Идентификатор клиента веб-приложения надстройки для MSA, зарегистрированного в msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="dc753-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="dc753-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="dc753-122">**Resource**</span></span>  |  <span data-ttu-id="dc753-123">Да</span><span class="sxs-lookup"><span data-stu-id="dc753-123">Yes</span></span>   |  <span data-ttu-id="dc753-124">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="dc753-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="dc753-125">Scopes</span><span class="sxs-lookup"><span data-stu-id="dc753-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="dc753-126">Да</span><span class="sxs-lookup"><span data-stu-id="dc753-126">Yes</span></span>  |  <span data-ttu-id="dc753-127">Задает разрешения, необходимые надстройке для ресурса, например Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="dc753-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="dc753-128">Authorizations</span><span class="sxs-lookup"><span data-stu-id="dc753-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="dc753-129">Нет</span><span class="sxs-lookup"><span data-stu-id="dc753-129">No</span></span>   | <span data-ttu-id="dc753-130">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="dc753-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="dc753-131">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="dc753-131">WebApplicationInfo example</span></span>

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
