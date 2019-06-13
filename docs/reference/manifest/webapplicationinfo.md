---
title: Элемент WebApplicationInfo в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bdd327f942009e255dd2515fb926d294212ecec8
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910322"
---
# <a name="webapplicationinfo-element"></a>Элемент WebApplicationInfo

Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:

- *Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.
- *Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.

> [!NOTE]
> В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint. Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets). Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.  

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Id**    |  Да   |  **Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  **Resource**  |  Да   |  Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  [Scopes](scopes.md)                |  Нет  |  Указывает разрешения, необходимые надстройке для работы с Microsoft Graph.  |

> [!NOTE] 
> В настоящее время необходимо, чтобы ресурс надстройки соответствовал ее ведущему приложению. Office запрашивает маркер для надстройки, только если может подтвердить право собственности. В настоящее время для этого необходимо, чтобы надстройка размещалась под полным доменным именем ресурса.

## <a name="webapplicationinfo-example"></a>Пример WebApplicationInfo

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
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
