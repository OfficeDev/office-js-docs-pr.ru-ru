---
title: Элемент Scopes в файле манифеста
description: Элемент Scopes содержит разрешения, необходимые для подключения надстройки к внешнему ресурсу.
ms.date: 08/12/2019
ms.localizationpriority: medium
ms.openlocfilehash: 346a143fdba35a153229b00052a463f726fd9056
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154841"
---
# <a name="scopes-element"></a>Элемент Scopes

Содержит разрешения, необходимые надстройки для внешнего ресурса, например Microsoft Graph. Если microsoft Graph является ресурсом, AppSource использует элемент Scopes для создания диалоговое окно согласия. Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.

**Области —** это детский элемент элементов [WebApplicationInfo](webapplicationinfo.md) и [Authorization](authorization.md) в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Scope**                |  Да     |   Имя разрешения; например, Files.Read.All или профиль. |

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
