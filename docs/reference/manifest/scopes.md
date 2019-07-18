---
title: Элемент Scopes в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdc9ebeb6fe4167a5ed5e9407f6ecc82d5b8d507
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771788"
---
# <a name="scopes-element"></a><span data-ttu-id="ae03a-102">Элемент Scopes</span><span class="sxs-lookup"><span data-stu-id="ae03a-102">Scopes element</span></span>

<span data-ttu-id="ae03a-103">Содержит разрешения, необходимые надстройке для работы с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="ae03a-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="ae03a-104">AppSource использует элемент scopes для создания диалогового окна согласия.</span><span class="sxs-lookup"><span data-stu-id="ae03a-104">AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="ae03a-105">Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="ae03a-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ae03a-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ae03a-106">Child elements</span></span>

|  <span data-ttu-id="ae03a-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="ae03a-107">Element</span></span> |  <span data-ttu-id="ae03a-108">Тип</span><span class="sxs-lookup"><span data-stu-id="ae03a-108">Type</span></span>  |  <span data-ttu-id="ae03a-109">Описание</span><span class="sxs-lookup"><span data-stu-id="ae03a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ae03a-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="ae03a-110">**Scope**</span></span>                |  <span data-ttu-id="ae03a-111">string</span><span class="sxs-lookup"><span data-stu-id="ae03a-111">string</span></span>     |   <span data-ttu-id="ae03a-112">Имя разрешения на доступ к Microsoft Graph (например, Files.Read.All).</span><span class="sxs-lookup"><span data-stu-id="ae03a-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="ae03a-113">Пример</span><span class="sxs-lookup"><span data-stu-id="ae03a-113">Example</span></span>

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
