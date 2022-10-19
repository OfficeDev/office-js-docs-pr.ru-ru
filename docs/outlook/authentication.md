---
title: Варианты проверки подлинности в надстройках Outlook
description: Надстройки Outlook предоставляют несколько различных способов проверки подлинности для разных сценариев.
ms.date: 10/17/2022
ms.localizationpriority: high
ms.openlocfilehash: d8ae8971c4095e5314885514226cd8f52728fb07
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607529"
---
# <a name="authentication-options-in-outlook-add-ins"></a>Варианты проверки подлинности в надстройках Outlook

Надстройка Outlook может получать доступ к данным из любого расположения в Интернете с сервера, на котором размещается надстройка, из внутренней сети или из другого места в облаке. Если эти сведения защищены, то надстройке нужен способ проверки подлинности пользователя. Надстройки Outlook предоставляют несколько различных способов проверки подлинности для разных сценариев.

## <a name="single-sign-on-access-token"></a>Маркер доступа для единого входа

Маркеры доступа для единого входа упрощают проверку подлинности надстройки и получение маркеров доступа для вызова [API Microsoft Graph](/graph/overview). Эта возможность повышает удобство работы, так как пользователю не требуется вводить свои учетные данные.

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Рекомендуем использовать маркеры единого входа в таких случаях:

- В основном надстройку применяют пользователи Microsoft 365.
- Надстройке требуется доступ к следующему:
  - службам Майкрософт, предоставляемым в составе Microsoft Graph;
  - сторонней службе, которой управляете вы.

Метод проверки подлинности для единого входа использует [поток "от имени" OAuth2, предоставленный службой Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of). Для этого необходимо, чтобы надстройка была зарегистрирована на [портале регистрации приложений](https://apps.dev.microsoft.com/), а в ее манифесте были указаны необходимые области Microsoft Graph.

> [!NOTE]
> Если надстройка использует манифест Teams для надстроек [Office (](../develop/json-manifest-overview.md)предварительная версия), существует конфигурация манифеста, но области Microsoft Graph не указаны. Надстройки с поддержкой единого входа, использующие манифест Teams, могут быть загружены неопубликованно, но в данный момент не могут быть развернуты другим способом.

Так надстройка может получить маркер доступа, действующий во внутреннем API сервера. Надстройка использует этот маркер в качестве токена носителя в заголовке `Authorization`, чтобы выполнять проверку подлинности обратного вызова API. После этого сервер сможет следующее:

- выполнить поток "от имени", чтобы получить маркер доступа, действующий в API Microsoft Graph;
- использовать сведения об удостоверении в маркере для определения удостоверения пользователя и проверки подлинности во внутренних службах.

Дополнительные сведения см. в [полном обзоре способа проверки подлинности с единым входом](../develop/sso-in-office-add-ins.md).

Дополнительные сведения об использовании маркера единого входа в надстройке Outlook см. в статье [Проверка подлинности пользователя с помощью маркера единого входа в надстройке Outlook](authenticate-a-user-with-an-sso-token.md).

Пример надстройки, использующей маркер единого входа, см. в статье [Единый вход надстройки Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).

## <a name="exchange-user-identity-token"></a>Получение маркера удостоверения Exchange

Маркеры удостоверений пользователей Exchange позволяют надстройке определить удостоверение пользователя. После проверки удостоверения пользователя ваше решение сможет выполнить одноразовую проверку подлинности во внутренней системе, а затем — принять маркер удостоверения пользователя для авторизации при последующих запросах. Используйте маркер удостоверения пользователя Exchange:

- если надстройку в основном применяют локальные пользователи Exchange;
- если надстройке требуется доступ к управляемой вами сторонней службе;
- в качестве резервной проверки подлинности при запуске надстройки в версии Office, не поддерживающей единый вход.

Надстройка может вызывать метод [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getuseridentitytokenasync-member(1)), чтобы получать маркеры удостоверений пользователей Exchange. Сведения об использовании этих маркеров см. в статье [Проверка подлинности пользователя с помощью маркера удостоверения для Exchange](authenticate-a-user-with-an-identity-token.md).

## <a name="access-tokens-obtained-via-oauth2-flows"></a>Маркеры доступа, полученные через потоки OAuth2

Надстройки также могут получать доступ к службам Майкрософт и сторонних разработчиков, поддерживающим протокол OAuth2 для авторизации. Рекомендуем использовать маркеры OAuth2 в таких случаях:

- надстройке требуется доступ к службе, которой вы не управляете.

Если вы применяете этот способ, надстройка предлагает пользователю войти в службу, используя метод [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) для инициализации потока OAuth2.

## <a name="callback-tokens"></a>Маркеры обратного вызова

Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api). Consider using callback tokens if your add-in:

- надстройке требуется доступ к почтовому ящику пользователя на внутреннем сервере.

Надстройки получают маркеры обратного вызова с помощью одного из методов [getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods). Уровень доступа контролируется при помощи разрешений, указываемых в манифесте надстройки.
