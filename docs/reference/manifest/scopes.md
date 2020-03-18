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
# <a name="scopes-element"></a>Элемент Scopes

Содержит разрешения, необходимые надстройке для внешнего ресурса, например Microsoft Graph. Когда Microsoft Graph является ресурсом, AppSource использует элемент scopes для создания диалогового окна согласия. Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.

**Области** — это дочерний элемент элементов [WebApplicationInfo](webapplicationinfo.md) и [authorization](authorization.md) в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Scope**                |  Да     |   Имя разрешения; Например, Files. Read. ALL или Profile. |

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
