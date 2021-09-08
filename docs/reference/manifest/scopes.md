---
title: Элемент Scopes в файле манифеста
description: Элемент Scopes содержит разрешения, необходимые для подключения надстройки к внешнему ресурсу.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: be68033e86de736703d9d1593ad361918d5a147d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937781"
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
