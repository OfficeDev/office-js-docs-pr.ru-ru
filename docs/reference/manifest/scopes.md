---
title: Элемент Scopes в файле манифеста
description: Элемент Scopes содержит разрешения, необходимые для подключения надстройки к внешнему ресурсу.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 05582ae05c13fae8e2272de3fe6111c5ff639f938a817fd0b50ad22e4234d033
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087260"
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
