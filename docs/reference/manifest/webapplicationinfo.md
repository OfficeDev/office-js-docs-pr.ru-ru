---
title: Элемент WebApplicationInfo в файле манифеста
description: Справочная документация элемента WebApplicationInfo для Office файлов манифеста надстройок (XML).
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb21c584f516fc9e50bdd881a383fb03f01c753c
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681551"
---
# <a name="webapplicationinfo-element"></a>Элемент WebApplicationInfo

Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:

- Ресурс OAuth 2.0, на который Office клиентского приложения могут потребоваться разрешения. 
- *Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.

> [!NOTE]
> В настоящее время поддерживается единый API входов для Word, Excel, Outlook и PowerPoint. Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](../requirement-sets/identity-api-requirement-sets.md). Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Microsoft 365. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.  

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Id**    |  Да   |  **Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  **Resource**  |  Да   |  Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  [Scopes](scopes.md)                |  Да  |  Указывает разрешения, необходимые надстройки для ресурса, например Microsoft Graph.  |

## <a name="webapplicationinfo-example"></a>Пример WebApplicationInfo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc</Resource>
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
