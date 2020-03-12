---
title: Элемент WebApplicationInfo в файле манифеста
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b6cf82776f683929845df83c642b28ad024d665a
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596734"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="79766-102">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="79766-102">WebApplicationInfo element</span></span>

<span data-ttu-id="79766-103">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="79766-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="79766-104">*Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="79766-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="79766-105">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="79766-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="79766-106">В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="79766-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="79766-107">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="79766-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="79766-108">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="79766-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="79766-109">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="79766-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="79766-110">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="79766-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="79766-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="79766-111">Child elements</span></span>

|  <span data-ttu-id="79766-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="79766-112">Element</span></span> |  <span data-ttu-id="79766-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="79766-113">Required</span></span>  |  <span data-ttu-id="79766-114">Описание</span><span class="sxs-lookup"><span data-stu-id="79766-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="79766-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="79766-115">**Id**</span></span>    |  <span data-ttu-id="79766-116">Да</span><span class="sxs-lookup"><span data-stu-id="79766-116">Yes</span></span>   |  <span data-ttu-id="79766-117">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="79766-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="79766-118">**мсаид**</span><span class="sxs-lookup"><span data-stu-id="79766-118">**MsaId**</span></span>    |  <span data-ttu-id="79766-119">Нет</span><span class="sxs-lookup"><span data-stu-id="79766-119">No</span></span>   |  <span data-ttu-id="79766-120">Идентификатор клиента веб-приложения надстройки для MSA, зарегистрированного в msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="79766-120">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="79766-121">**Resource**</span><span class="sxs-lookup"><span data-stu-id="79766-121">**Resource**</span></span>  |  <span data-ttu-id="79766-122">Да</span><span class="sxs-lookup"><span data-stu-id="79766-122">Yes</span></span>   |  <span data-ttu-id="79766-123">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="79766-123">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="79766-124">Scopes</span><span class="sxs-lookup"><span data-stu-id="79766-124">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="79766-125">Да</span><span class="sxs-lookup"><span data-stu-id="79766-125">Yes</span></span>  |  <span data-ttu-id="79766-126">Задает разрешения, необходимые надстройке для ресурса, например Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="79766-126">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="79766-127">Authorizations</span><span class="sxs-lookup"><span data-stu-id="79766-127">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="79766-128">Нет</span><span class="sxs-lookup"><span data-stu-id="79766-128">No</span></span>   | <span data-ttu-id="79766-129">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="79766-129">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="79766-130">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="79766-130">WebApplicationInfo example</span></span>

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
