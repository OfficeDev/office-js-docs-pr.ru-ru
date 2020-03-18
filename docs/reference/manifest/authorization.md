---
title: Элемент authorization в файле манифеста
description: Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cece0934eb9db3175b173e97d7ab478827b7cda2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718442"
---
# <a name="authorization-element"></a>Элемент authorization

Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.

**Авторизация** является дочерним элементом элемента [authorizations](authorizations.md) в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Resource**  |  Да   |  Задает URL-адрес внешнего ресурса.|
|  [Scopes](scopes.md)                |  Да  |  Задает разрешения, необходимые надстройке для ресурса.  |

## <a name="example"></a>Пример

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
