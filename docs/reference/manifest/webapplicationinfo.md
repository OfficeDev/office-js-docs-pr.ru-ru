---
title: Элемент WebApplicationInfo в файле манифеста
description: Справочная документация по элементу VersionOverrides для файлов манифеста надстроек Office (XML).
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6acd0d5688bdd93d4054d0589afe5517afb1296f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720297"
---
# <a name="webapplicationinfo-element"></a>Элемент WebApplicationInfo

Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:

- *Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.
- *Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.

> [!NOTE]
> В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint. Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](../requirement-sets/identity-api-requirement-sets.md). Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.  

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Id**    |  Да   |  **Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  **мсаид**    |  Нет   |  Идентификатор клиента веб-приложения надстройки для MSA, зарегистрированного в msm.live.com.|
|  **Resource**  |  Да   |  Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  [Scopes](scopes.md)                |  Да  |  Задает разрешения, необходимые надстройке для ресурса, например Microsoft Graph.  |
|  [Authorizations](authorizations.md)  |  Нет   | Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.|

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
