---
title: Элемент authorization в файле манифеста
description: Указывает внешний ресурс, на который веб-приложению надстройки требуется авторизация и необходимые разрешения.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294264"
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
