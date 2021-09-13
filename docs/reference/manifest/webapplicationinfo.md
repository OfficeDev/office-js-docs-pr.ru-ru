---
title: Элемент WebApplicationInfo в файле манифеста
description: Справочная документация элемента WebApplicationInfo для Office файлов манифеста надстройок (XML).
ms.date: 07/30/2020
ms.localizationpriority: medium
ms.openlocfilehash: 7de9271fc3e7ed76c0423c8a0b8ab70360b105c3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154548"
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
|  **MsaId**    |  Нет   |  ID клиента веб-приложения надстройки для MSA, зарегистрированного в msm.live.com.|
|  **Resource**  |  Да   |  Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  [Scopes](scopes.md)                |  Да  |  Указывает разрешения, необходимые надстройки для ресурса, например Microsoft Graph.  |
|  [Authorizations](authorizations.md)  |  Нет   | Указывает внешние ресурсы, на которые веб-приложению надстройки требуется авторизация, и необходимые разрешения.|

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
