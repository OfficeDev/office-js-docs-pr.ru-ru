---
title: Элемент авторизации в файле манифеста
description: Указывает внешний ресурс, на который веб-приложению надстройки требуется авторизация и необходимые разрешения.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936733"
---
# <a name="authorization-element"></a>Элемент авторизации

Указывает внешние ресурсы, на которые веб-приложению надстройки требуется авторизация, и необходимые разрешения.

**Авторизация** — это детский элемент элемента [Авторизация](authorizations.md) в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Resource**  |  Да   |  Указывает URL-адрес внешнего ресурса.|
|  [Scopes](scopes.md)                |  Да  |  Указывает разрешения, необходимые надстройки для ресурса.  |

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
