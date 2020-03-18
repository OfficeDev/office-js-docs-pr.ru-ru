---
title: Элемент WebApplicationInfo в файле манифеста
description: Справочная документация по элементу VersionOverrides для файлов манифеста надстроек Office (XML).
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6acd0d5688bdd93d4054d0589afe5517afb1296f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720297"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="de2c6-103">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="de2c6-103">WebApplicationInfo element</span></span>

<span data-ttu-id="de2c6-104">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="de2c6-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="de2c6-105">*Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="de2c6-105">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="de2c6-106">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="de2c6-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="de2c6-107">В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="de2c6-107">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="de2c6-108">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="de2c6-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="de2c6-109">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="de2c6-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="de2c6-110">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="de2c6-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="de2c6-111">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="de2c6-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="de2c6-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="de2c6-112">Child elements</span></span>

|  <span data-ttu-id="de2c6-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="de2c6-113">Element</span></span> |  <span data-ttu-id="de2c6-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="de2c6-114">Required</span></span>  |  <span data-ttu-id="de2c6-115">Описание</span><span class="sxs-lookup"><span data-stu-id="de2c6-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="de2c6-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="de2c6-116">**Id**</span></span>    |  <span data-ttu-id="de2c6-117">Да</span><span class="sxs-lookup"><span data-stu-id="de2c6-117">Yes</span></span>   |  <span data-ttu-id="de2c6-118">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="de2c6-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="de2c6-119">**мсаид**</span><span class="sxs-lookup"><span data-stu-id="de2c6-119">**MsaId**</span></span>    |  <span data-ttu-id="de2c6-120">Нет</span><span class="sxs-lookup"><span data-stu-id="de2c6-120">No</span></span>   |  <span data-ttu-id="de2c6-121">Идентификатор клиента веб-приложения надстройки для MSA, зарегистрированного в msm.live.com.</span><span class="sxs-lookup"><span data-stu-id="de2c6-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="de2c6-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="de2c6-122">**Resource**</span></span>  |  <span data-ttu-id="de2c6-123">Да</span><span class="sxs-lookup"><span data-stu-id="de2c6-123">Yes</span></span>   |  <span data-ttu-id="de2c6-124">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="de2c6-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="de2c6-125">Scopes</span><span class="sxs-lookup"><span data-stu-id="de2c6-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="de2c6-126">Да</span><span class="sxs-lookup"><span data-stu-id="de2c6-126">Yes</span></span>  |  <span data-ttu-id="de2c6-127">Задает разрешения, необходимые надстройке для ресурса, например Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="de2c6-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="de2c6-128">Authorizations</span><span class="sxs-lookup"><span data-stu-id="de2c6-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="de2c6-129">Нет</span><span class="sxs-lookup"><span data-stu-id="de2c6-129">No</span></span>   | <span data-ttu-id="de2c6-130">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="de2c6-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="de2c6-131">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="de2c6-131">WebApplicationInfo example</span></span>

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
