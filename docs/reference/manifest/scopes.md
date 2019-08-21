---
title: Элемент Scopes в файле манифеста
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 1e36bdcd0cdcaa8c842e924c2543d56bdc4e26a7
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477735"
---
# <a name="scopes-element"></a><span data-ttu-id="50698-102">Элемент Scopes</span><span class="sxs-lookup"><span data-stu-id="50698-102">Scopes element</span></span>

<span data-ttu-id="50698-103">Содержит разрешения, необходимые надстройке для внешнего ресурса, например Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="50698-103">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="50698-104">Когда Microsoft Graph является ресурсом, AppSource использует элемент scopes для создания диалогового окна согласия.</span><span class="sxs-lookup"><span data-stu-id="50698-104">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="50698-105">Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="50698-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="50698-106">**Области** — это дочерний элемент элементов [WebApplicationInfo](webapplicationinfo.md) и [authorization](authorization.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="50698-106">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="50698-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="50698-107">Child elements</span></span>

|  <span data-ttu-id="50698-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="50698-108">Element</span></span> |  <span data-ttu-id="50698-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="50698-109">Required</span></span>  |  <span data-ttu-id="50698-110">Описание</span><span class="sxs-lookup"><span data-stu-id="50698-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="50698-111">**Scope**</span><span class="sxs-lookup"><span data-stu-id="50698-111">**Scope**</span></span>                |  <span data-ttu-id="50698-112">Да</span><span class="sxs-lookup"><span data-stu-id="50698-112">Yes</span></span>     |   <span data-ttu-id="50698-113">Имя разрешения; Например, Files. Read. ALL или Profile.</span><span class="sxs-lookup"><span data-stu-id="50698-113">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="50698-114">Пример</span><span class="sxs-lookup"><span data-stu-id="50698-114">Example</span></span>

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
