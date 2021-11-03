---
title: Элемент Scopes в файле манифеста
description: Элемент Scopes содержит разрешения, необходимые для подключения надстройки к внешнему ресурсу.
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 16e8a19a7aa73efa6aac00c915fde8d2b8647bad
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681537"
---
# <a name="scopes-element"></a>Элемент Scopes

Содержит разрешения, необходимые надстройки для внешнего ресурса, например Microsoft Graph. Если microsoft Graph является ресурсом, AppSource использует элемент Scopes для создания диалоговое окно согласия. Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.

**Области —** это детский элемент элемента [WebApplicationInfo](webapplicationinfo.md) в манифесте.

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
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
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
