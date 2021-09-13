---
title: Элемент Авторизации в файле манифеста
description: Указывает внешние ресурсы, на которые веб-приложению надстройки требуется авторизация, и необходимые разрешения.
ms.date: 08/12/2019
ms.localizationpriority: medium
ms.openlocfilehash: 4b13e26f13fae6fefd579868df8b67dd94cb35c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153721"
---
# <a name="authorizations-element"></a>Элемент Авторизация

Указывает внешние ресурсы, на которые веб-приложению надстройки требуется авторизация, и необходимые разрешения.

**Авторизация** — это детский элемент элемента [WebApplicationInfo](webapplicationinfo.md) в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Authorization](authorization.md)                |  Да     |   Определяет внешний ресурс, на который веб-приложению надстройки требуется авторизация, и необходимые ему области (разрешения). |

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
