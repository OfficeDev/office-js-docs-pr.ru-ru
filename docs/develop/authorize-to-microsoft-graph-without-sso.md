---
title: Авторизация в Microsoft Graph без единого входа
description: Узнайте, как осуществлять авторизацию в Microsoft Graph без единого входа
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4f96c65fcc3c90a616f43189e1facebdbf8e9a8c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150742"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Авторизация в Microsoft Graph без единого входа

Ваша надстройка может получить авторизацию в Microsoft Graph данных, получив маркер доступа к Microsoft Graph из Azure Active Directory Azure AD. Используйте поток кода авторизации или неявный поток так же, как и в других веб-приложениях, но за одним исключением: Azure AD не позволяет его входу на страницу открываться в iframe. При работе с надстройкой Office в *Office в Интернете* область задач является элементом iframe. Это означает, что вам потребуется открыть экран входа Azure AD в диалоговом окне, открытом Office диалоговом API. Это повлияет на способ использования библиотек помощника проверки подлинности и авторизации. Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).

Сведения о проверке подлинности в Azure AD см. в обзоре [платформа удостоверений Майкрософт (v2.0),](/azure/active-directory/develop/v2-overview)в котором вы найдете руководства и руководства в этом наборе документации, а также ссылки на соответствующие примеры. Опять же, вам может потребоваться изменить код из примеров, чтобы запустить его в диалоговом окне Office, учитывая, что диалоговое окно Office запускается в отдельном процессе из области задач.

После получения кодом маркера доступа в Microsoft Graph либо он передает маркер доступа из диалогового окна в области задач, либо сохраняет маркер в базе данных и сигнализирует области задач о том, что маркер доступен. (Подробные сведения см. в [Office диалоговом API.)](auth-with-office-dialog-api.md) Код в области задач запрашивает данные из Microsoft Graph и включает маркер в эти запросы. Дополнительные сведения о вызове microsoft Graph и SDKs microsoft Graph см. в документации [Microsoft Graph.](/graph/)

## <a name="recommended-libraries-and-samples"></a>Рекомендуемые библиотеки и примеры

Мы рекомендуем использовать следующие библиотеки при доступе к microsoft Graph без использования SSO.

- Для надстроек, использующих серверную часть с платформой на основе .NET, например .NET Core или ASP.NET, используйте [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Для надстроек, использующих серверную часть на основе NodeJS, используйте [Azure AD Passport](https://github.com/AzureAD/passport-azure-ad).
- Для надстроек, использующих неявный поток, используйте [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Дополнительные сведения о рекомендуемых библиотеках для работы с платформой удостоверений Майкрософт (ранее — AAD версии 2.0) см. в статье [Библиотеки проверки подлинности платформы удостоверений Майкрософт](/azure/active-directory/develop/reference-v2-libraries).

В следующих примерах Graph microsoft Office надстройки.

- [Надстройка Office в Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Outlook в Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Office в Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
