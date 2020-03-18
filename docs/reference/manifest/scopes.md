---
title: Элемент Scopes в файле манифеста
description: Элемент scopes содержит разрешения, необходимые надстройке для подключения к внешнему ресурсу.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 69a394b4cbe324b49c03425e6b2df92f44cbd21f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717924"
---
# <a name="scopes-element"></a><span data-ttu-id="3df47-103">Элемент Scopes</span><span class="sxs-lookup"><span data-stu-id="3df47-103">Scopes element</span></span>

<span data-ttu-id="3df47-104">Содержит разрешения, необходимые надстройке для внешнего ресурса, например Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3df47-104">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="3df47-105">Когда Microsoft Graph является ресурсом, AppSource использует элемент scopes для создания диалогового окна согласия.</span><span class="sxs-lookup"><span data-stu-id="3df47-105">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="3df47-106">Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3df47-106">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="3df47-107">**Области** — это дочерний элемент элементов [WebApplicationInfo](webapplicationinfo.md) и [authorization](authorization.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="3df47-107">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="3df47-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="3df47-108">Child elements</span></span>

|  <span data-ttu-id="3df47-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="3df47-109">Element</span></span> |  <span data-ttu-id="3df47-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="3df47-110">Required</span></span>  |  <span data-ttu-id="3df47-111">Описание</span><span class="sxs-lookup"><span data-stu-id="3df47-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3df47-112">**Scope**</span><span class="sxs-lookup"><span data-stu-id="3df47-112">**Scope**</span></span>                |  <span data-ttu-id="3df47-113">Да</span><span class="sxs-lookup"><span data-stu-id="3df47-113">Yes</span></span>     |   <span data-ttu-id="3df47-114">Имя разрешения; Например, Files. Read. ALL или Profile.</span><span class="sxs-lookup"><span data-stu-id="3df47-114">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="3df47-115">Пример</span><span class="sxs-lookup"><span data-stu-id="3df47-115">Example</span></span>

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
