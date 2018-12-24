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
# <a name="scopes-element"></a>Элемент Scopes

Содержит разрешения, необходимые надстройке для работы с Microsoft Graph. Магазин Office использует элемент Scopes для создания диалогового окна подтверждения. Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Тип  |  Описание  |
|:-----|:-----|:-----|
|  **Scope**                |  string     |   Имя разрешения на доступ к Microsoft Graph (например, Files.Read.All). |

## <a name="example"></a>Пример

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
