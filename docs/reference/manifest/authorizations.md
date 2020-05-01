---
title: Элемент authorizations в файле манифеста
description: Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 7ae0b9d0ec32a20846142a9fc89c48fe9cdf8053
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720661"
---
# <a name="authorizations-element"></a><span data-ttu-id="778a1-103">Элемент authorizations</span><span class="sxs-lookup"><span data-stu-id="778a1-103">Authorizations element</span></span>

<span data-ttu-id="778a1-104">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="778a1-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="778a1-105">**Авторизация** является дочерним элементом элемента [WebApplicationInfo](webapplicationinfo.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="778a1-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="778a1-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="778a1-106">Child elements</span></span>

|  <span data-ttu-id="778a1-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="778a1-107">Element</span></span> |  <span data-ttu-id="778a1-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="778a1-108">Required</span></span>  |  <span data-ttu-id="778a1-109">Описание</span><span class="sxs-lookup"><span data-stu-id="778a1-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="778a1-110">Авторизация</span><span class="sxs-lookup"><span data-stu-id="778a1-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="778a1-111">Да</span><span class="sxs-lookup"><span data-stu-id="778a1-111">Yes</span></span>     |   <span data-ttu-id="778a1-112">Определяет внешний ресурс, на который веб-приложение надстройки должно выполнять авторизацию, и необходимые области (разрешения).</span><span class="sxs-lookup"><span data-stu-id="778a1-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="778a1-113">Пример</span><span class="sxs-lookup"><span data-stu-id="778a1-113">Example</span></span>

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
