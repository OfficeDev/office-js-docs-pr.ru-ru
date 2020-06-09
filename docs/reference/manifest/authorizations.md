---
title: Элемент authorizations в файле манифеста
description: Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 675585f99fc6261a2145219d553f02b9f9abded3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608756"
---
# <a name="authorizations-element"></a><span data-ttu-id="0cc3c-103">Элемент authorizations</span><span class="sxs-lookup"><span data-stu-id="0cc3c-103">Authorizations element</span></span>

<span data-ttu-id="0cc3c-104">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="0cc3c-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="0cc3c-105">**Авторизация** является дочерним элементом элемента [WebApplicationInfo](webapplicationinfo.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="0cc3c-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0cc3c-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="0cc3c-106">Child elements</span></span>

|  <span data-ttu-id="0cc3c-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="0cc3c-107">Element</span></span> |  <span data-ttu-id="0cc3c-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0cc3c-108">Required</span></span>  |  <span data-ttu-id="0cc3c-109">Описание</span><span class="sxs-lookup"><span data-stu-id="0cc3c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0cc3c-110">Authorization</span><span class="sxs-lookup"><span data-stu-id="0cc3c-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="0cc3c-111">Да</span><span class="sxs-lookup"><span data-stu-id="0cc3c-111">Yes</span></span>     |   <span data-ttu-id="0cc3c-112">Определяет внешний ресурс, на который веб-приложение надстройки должно выполнять авторизацию, и необходимые области (разрешения).</span><span class="sxs-lookup"><span data-stu-id="0cc3c-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="0cc3c-113">Пример</span><span class="sxs-lookup"><span data-stu-id="0cc3c-113">Example</span></span>

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
