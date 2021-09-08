---
title: Проверка подлинности пользователя с помощью маркера единого входа
description: Узнайте, как реализовать единый вход в службе с помощью маркера единого входа, предоставляемого надстройкой Outlook.
ms.date: 09/03/2021
localization_priority: Normal
ms.openlocfilehash: a30d5671286f9a0afcff6fd07d4fcfcb4c830c36
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938143"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>Проверка подлинности пользователя с помощью маркера с одним входом в Outlook надстройки

Единый вход (SSO) упрощает проверку подлинности пользователей в надстройке (и, при необходимости, получение маркеров доступа для вызова [API Microsoft Graph](/graph/overview)).

Так надстройка может получить маркер доступа, действующий во внутреннем API сервера. Надстройка использует этот маркер в качестве токена носителя в заголовке `Authorization`, чтобы выполнять проверку подлинности обратного вызова API. Кроме того, можно использовать код на стороне сервера.

- выполнить поток "от имени", чтобы получить маркер доступа, действующий в API Microsoft Graph;
- использовать сведения об удостоверении в маркере для определения удостоверения пользователя и проверки подлинности во внутренних службах.

Общие сведения о едином входе в надстройках Office см. в статье [Включение единого входа для надстроек Office (тестовый режим)](../develop/sso-in-office-add-ins.md) и [Авторизация для Microsoft Graph в надстройке Office](../develop/authorize-to-microsoft-graph.md).

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Включить современную проверку подлинности в Microsoft 365 аренды

Чтобы использовать SSO с Outlook надстройки, необходимо включить современную проверку подлинности для Microsoft 365 аренды. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="register-your-add-in"></a>Регистрация надстройки

Чтобы использовать единый вход, надстройке Outlook потребуется серверный веб-API, зарегистрированный в Azure Active Directory (AAD) версии 2.0. Дополнительные сведения см. в статье [Регистрация надстройки Office, использующей единый вход с конечной точкой Microsoft Azure AD версии 2.0](../develop/register-sso-add-in-aad-v2.md).

### <a name="provide-consent-when-sideloading-an-add-in"></a>Предоставление согласия при загрузке неопубликованной надстройки

При разработке надстройки необходимо заранее предоставить согласие. Дополнительные сведения см. в [записи согласия администратора гранта на надстройки.](../develop/grant-admin-consent-to-an-add-in.md)

## <a name="update-the-add-in-manifest"></a>Обновление манифеста надстройки

Следующий этап включения единого входа в надстройке — добавление элемента `WebApplicationInfo` в конце элемента [VersionOverrides](../reference/manifest/versionoverrides.md) библиотеки `VersionOverridesV1_1`. Дополнительные сведения см. в статье [Конфигурация надстройки](../develop/sso-in-office-add-ins.md#configure-the-add-in).

## <a name="get-the-sso-token"></a>Получение маркера единого входа

Надстройка получает маркер единого входа с помощью клиентского скрипта. Дополнительные сведения см. в разделе [Добавление кода для клиента](../develop/sso-in-office-add-ins.md#add-client-side-code).

## <a name="use-the-sso-token-at-the-back-end"></a>Использование маркера единого входа во внутренней службе

В большинстве случаев практически нет смысла получать маркер доступа, если надстройка не передает его на сторону сервера и не использует его там. Дополнительные сведения о том, какие действия должны выполняться на стороне сервера, а какие нет, см. в разделе [Добавление серверного кода](../develop/sso-in-office-add-ins.md#add-server-side-code).

> [!IMPORTANT]
> При использовании маркера единого входа в качестве удостоверения в *Outlook* рекомендуем также [использовать маркер удостоверения Exchange](authenticate-a-user-with-an-identity-token.md) в качестве альтернативного удостоверения. Пользователи надстройки могут использовать различные клиенты, не все и которых поддерживают предоставление маркера единого входа. Если в качестве альтернативы используется маркер удостоверения Exchange, вы можете избежать повторного запрашивания учетных данных этих пользователей. Дополнительные сведения см. в статье [Сценарий: реализация единого входа для службы в надстройке Outlook](implement-sso-in-outlook-add-in.md).

## <a name="see-also"></a>См. также

- Пример надстройки Outlook которая использует маркер SSO для доступа к API Microsoft Graph, см. в Outlook [SSO надстройки.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Справочные материалы по API единого входа](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [Настройка требования IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md)
