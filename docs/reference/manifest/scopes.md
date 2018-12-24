---
title: Элемент Scopes в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 01d34481b14ac6a9186de07d352b9985dc1695a4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432643"
---
# <a name="scopes-element"></a><span data-ttu-id="887b9-102">Элемент Scopes</span><span class="sxs-lookup"><span data-stu-id="887b9-102">Scopes element</span></span>

<span data-ttu-id="887b9-103">Содержит разрешения, необходимые надстройке для работы с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="887b9-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="887b9-104">Магазин Office использует элемент Scopes для создания диалогового окна подтверждения.</span><span class="sxs-lookup"><span data-stu-id="887b9-104">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="887b9-105">Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="887b9-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="887b9-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="887b9-106">Child elements</span></span>

|  <span data-ttu-id="887b9-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="887b9-107">Element</span></span> |  <span data-ttu-id="887b9-108">Тип</span><span class="sxs-lookup"><span data-stu-id="887b9-108">Type</span></span>  |  <span data-ttu-id="887b9-109">Описание</span><span class="sxs-lookup"><span data-stu-id="887b9-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="887b9-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="887b9-110">**Scope**</span></span>                |  <span data-ttu-id="887b9-111">string</span><span class="sxs-lookup"><span data-stu-id="887b9-111">string</span></span>     |   <span data-ttu-id="887b9-112">Имя разрешения на доступ к Microsoft Graph (например, Files.Read.All).</span><span class="sxs-lookup"><span data-stu-id="887b9-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="887b9-113">Пример</span><span class="sxs-lookup"><span data-stu-id="887b9-113">Example</span></span>

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
