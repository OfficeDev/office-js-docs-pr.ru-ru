---
title: Элемент Scopes в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 903f7ff68313549234c07926cc63dc7e783ae400
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451943"
---
# <a name="scopes-element"></a><span data-ttu-id="96c7c-102">Элемент Scopes</span><span class="sxs-lookup"><span data-stu-id="96c7c-102">Scopes element</span></span>

<span data-ttu-id="96c7c-103">Содержит разрешения, необходимые надстройке для работы с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96c7c-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="96c7c-104">Магазин Office использует элемент Scopes для создания диалогового окна подтверждения.</span><span class="sxs-lookup"><span data-stu-id="96c7c-104">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="96c7c-105">Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="96c7c-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="96c7c-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="96c7c-106">Child elements</span></span>

|  <span data-ttu-id="96c7c-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="96c7c-107">Element</span></span> |  <span data-ttu-id="96c7c-108">Тип</span><span class="sxs-lookup"><span data-stu-id="96c7c-108">Type</span></span>  |  <span data-ttu-id="96c7c-109">Описание</span><span class="sxs-lookup"><span data-stu-id="96c7c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="96c7c-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="96c7c-110">**Scope**</span></span>                |  <span data-ttu-id="96c7c-111">string</span><span class="sxs-lookup"><span data-stu-id="96c7c-111">string</span></span>     |   <span data-ttu-id="96c7c-112">Имя разрешения на доступ к Microsoft Graph (например, Files.Read.All).</span><span class="sxs-lookup"><span data-stu-id="96c7c-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="96c7c-113">Пример</span><span class="sxs-lookup"><span data-stu-id="96c7c-113">Example</span></span>

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
