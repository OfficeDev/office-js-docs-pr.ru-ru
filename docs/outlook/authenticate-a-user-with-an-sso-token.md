---
title: Проверка подлинности пользователя с помощью маркера единого входа
description: Узнайте, как реализовать единый вход в службе с помощью маркера единого входа, предоставляемого надстройкой Outlook.
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: e0925979d26f6b3145658d71b1edaf30431e0c7e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293984"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>Проверка подлинности пользователя с помощью маркера единого входа в надстройке Outlook

Единый вход (SSO) упрощает проверку подлинности пользователей в надстройке (и, при необходимости, получение маркеров доступа для вызова [API Microsoft Graph](/graph/overview)).

Так надстройка может получить маркер доступа, действующий во внутреннем API сервера. Надстройка использует этот маркер в качестве токена носителя в заголовке `Authorization`, чтобы выполнять проверку подлинности обратного вызова API. При необходимости серверный код также может:

- выполнить поток "от имени", чтобы получить маркер доступа, действующий в API Microsoft Graph;
- использовать сведения об удостоверении в маркере для определения удостоверения пользователя и проверки подлинности во внутренних службах.

Общие сведения о едином входе в надстройках Office см. в статье [Включение единого входа для надстроек Office (тестовый режим)](../develop/sso-in-office-add-ins.md) и [Авторизация для Microsoft Graph в надстройке Office](../develop/authorize-to-microsoft-graph.md).

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Включение современной проверки подлинности в клиенте Microsoft 365

Чтобы использовать единый вход с надстройкой Outlook, необходимо включить современный способ проверки подлинности для клиента Microsoft 365. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="register-your-add-in"></a>Регистрация надстройки

Чтобы использовать единый вход, надстройке Outlook потребуется серверный веб-API, зарегистрированный в Azure Active Directory (AAD) версии 2.0. Дополнительные сведения см. в статье [Регистрация надстройки Office, использующей единый вход с конечной точкой Microsoft Azure AD версии 2.0](../develop/register-sso-add-in-aad-v2.md).

### <a name="provide-consent-when-sideloading-an-add-in"></a>Предоставление согласия при загрузке неопубликованной надстройки

Когда вы разрабатываете надстройку, вам потребуется предварительно предоставить согласие. Для получения дополнительных сведений обратитесь к разделу [Предоставление администратору разрешения для надстройки](../develop/grant-admin-consent-to-an-add-in.md).

## <a name="update-the-add-in-manifest"></a>Обновление манифеста надстройки

Следующий этап включения единого входа в надстройке — добавление элемента `WebApplicationInfo` в конце элемента [VersionOverrides](../reference/manifest/versionoverrides.md) библиотеки `VersionOverridesV1_1`. Дополнительные сведения см. в статье [Конфигурация надстройки](../develop/sso-in-office-add-ins.md#configure-the-add-in).

## <a name="get-the-sso-token"></a>Получение маркера единого входа

Надстройка получает маркер единого входа с помощью клиентского скрипта. Дополнительные сведения см. в разделе [Добавление кода для клиента](../develop/sso-in-office-add-ins.md#add-client-side-code).

## <a name="use-the-sso-token-at-the-back-end"></a>Использование маркера единого входа во внутренней службе

В большинстве случаев практически нет смысла получать маркер доступа, если надстройка не передает его на сторону сервера и не использует его там. Дополнительные сведения о том, какие действия должны выполняться на стороне сервера, а какие нет, см. в разделе [Добавление серверного кода](../develop/sso-in-office-add-ins.md#add-server-side-code).

> [!IMPORTANT]
> При использовании маркера единого входа в качестве удостоверения в *Outlook* рекомендуем также [использовать маркер удостоверения Exchange](authenticate-a-user-with-an-identity-token.md) в качестве альтернативного удостоверения. Пользователи надстройки могут использовать различные клиенты, не все и которых поддерживают предоставление маркера единого входа. Если в качестве альтернативы используется маркер удостоверения Exchange, вы можете избежать повторного запрашивания учетных данных этих пользователей. Дополнительные сведения см. в статье [Сценарий: реализация единого входа для службы в надстройке Outlook](implement-sso-in-outlook-add-in.md).

## <a name="see-also"></a>См. также

- Для примера надстройки Outlook, использующей маркер единого входа для доступа к API Microsoft Graph, обратитесь к разделу [единый вход надстройки Outlook](https://github.com/OfficeDev/Outlook-Add-in-SSO).
- [Справочные материалы по API единого входа](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [Настройка требования IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md)
