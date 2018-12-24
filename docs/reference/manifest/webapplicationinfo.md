---
title: Элемент WebApplicationInfo в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 14c375898a05aa957d2c2b906ebaba961c4f1c2d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432734"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="e05b0-102">Элемент WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="e05b0-102">WebApplicationInfo element</span></span>

<span data-ttu-id="e05b0-103">Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:</span><span class="sxs-lookup"><span data-stu-id="e05b0-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="e05b0-104">*Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="e05b0-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="e05b0-105">*Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="e05b0-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="e05b0-106">В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="e05b0-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="e05b0-107">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="e05b0-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span> <span data-ttu-id="e05b0-108">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="e05b0-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="e05b0-109">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="e05b0-109">To learn how to do this, see�[Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="e05b0-110">**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="e05b0-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="e05b0-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e05b0-111">Child elements</span></span>

|  <span data-ttu-id="e05b0-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="e05b0-112">Element</span></span> |  <span data-ttu-id="e05b0-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e05b0-113">Required</span></span>  |  <span data-ttu-id="e05b0-114">Описание</span><span class="sxs-lookup"><span data-stu-id="e05b0-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e05b0-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="e05b0-115">**Id**</span></span>    |  <span data-ttu-id="e05b0-116">Да</span><span class="sxs-lookup"><span data-stu-id="e05b0-116">Yes</span></span>   |  <span data-ttu-id="e05b0-117">**Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="e05b0-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="e05b0-118">**Resource**</span><span class="sxs-lookup"><span data-stu-id="e05b0-118">**Resource**</span></span>  |  <span data-ttu-id="e05b0-119">Да</span><span class="sxs-lookup"><span data-stu-id="e05b0-119">Yes</span></span>   |  <span data-ttu-id="e05b0-120">Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="e05b0-120">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="e05b0-121">Scopes</span><span class="sxs-lookup"><span data-stu-id="e05b0-121">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="e05b0-122">Нет</span><span class="sxs-lookup"><span data-stu-id="e05b0-122">No</span></span>  |  <span data-ttu-id="e05b0-123">Указывает разрешения, необходимые надстройке для работы с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="e05b0-123">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="e05b0-124">В настоящее время необходимо, чтобы ресурс надстройки соответствовал ее ведущему приложению.</span><span class="sxs-lookup"><span data-stu-id="e05b0-124">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="e05b0-125">Office запрашивает маркер для надстройки, только если может подтвердить право собственности. В настоящее время для этого необходимо, чтобы надстройка размещалась под полным доменным именем ресурса.</span><span class="sxs-lookup"><span data-stu-id="e05b0-125">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="e05b0-126">Пример WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="e05b0-126">WebApplicationInfo example</span></span>

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
