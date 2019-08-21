---
title: Элемент authorizations в файле манифеста
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6a271423ddd549431c2f580e2793faab3c49090e
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477959"
---
# <a name="authorizations-element"></a><span data-ttu-id="39ed5-102">Элемент authorizations</span><span class="sxs-lookup"><span data-stu-id="39ed5-102">Authorizations element</span></span>

<span data-ttu-id="39ed5-103">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="39ed5-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="39ed5-104">**Авторизация** является дочерним элементом элемента [WebApplicationInfo](webapplicationinfo.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="39ed5-104">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="39ed5-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="39ed5-105">Child elements</span></span>

|  <span data-ttu-id="39ed5-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="39ed5-106">Element</span></span> |  <span data-ttu-id="39ed5-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="39ed5-107">Required</span></span>  |  <span data-ttu-id="39ed5-108">Описание</span><span class="sxs-lookup"><span data-stu-id="39ed5-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="39ed5-109">Авторизация</span><span class="sxs-lookup"><span data-stu-id="39ed5-109">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="39ed5-110">Да</span><span class="sxs-lookup"><span data-stu-id="39ed5-110">Yes</span></span>     |   <span data-ttu-id="39ed5-111">Определяет внешний ресурс, на который веб-приложение надстройки должно выполнять авторизацию, и необходимые области (разрешения).</span><span class="sxs-lookup"><span data-stu-id="39ed5-111">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="39ed5-112">Пример</span><span class="sxs-lookup"><span data-stu-id="39ed5-112">Example</span></span>

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
