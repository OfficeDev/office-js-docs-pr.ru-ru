---
title: Авторизация в Microsoft Graph без единого входа
description: Узнайте, как осуществлять авторизацию в Microsoft Graph без единого входа
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 99d300d0155ba9a117efda5d31ef068a41eb86a9
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104835"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Авторизация в Microsoft Graph без единого входа

Ваша надстройка может получить авторизацию данных Microsoft Graph, получив маркер доступа к Microsoft Graph из Azure Active Directory (Azure AD). Используйте поток кода авторизации или неявный поток так же, как и в других веб-приложениях, но с одним исключением: Azure AD не разрешает открытие страницы входов в iframe. При работе с надстройкой Office в *Office в Интернете* область задач является элементом iframe. Это означает, что вам потребуется открыть экран входа в Azure AD в диалоговом окне, открываемом с помощью API диалоговых окне Office. Это повлияет на способ использования библиотек помощника проверки подлинности и авторизации. Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).

Сведения о программировании проверки подлинности с помощью Azure AD см. в обзоре платформы удостоверений [Майкрософт (v2.0),](/azure/active-directory/develop/v2-overview)где вы найдете учебники и руководства в этом наборе документации, а также ссылки на соответствующие примеры. Опять же, вам может потребоваться изменить код из примеров, чтобы запустить его в диалоговом окне Office, учитывая, что диалоговое окно Office запускается в отдельном процессе из области задач.

После того как код получит маркер доступа в Microsoft Graph, он либо передает маркер доступа из диалогового окна в области задач, либо сохраняет маркер в базе данных и сообщает области задач, что маркер доступен. (Подробные сведения см. в подразделе "Проверка подлинности с помощью [API диалогов](auth-with-office-dialog-api.md) Office".) Код в области задач запрашивает данные из Microsoft Graph и включает маркер в эти запросы. Дополнительные сведения о вызове Microsoft Graph и SDKs Microsoft Graph см. в [документации по Microsoft Graph.](/graph/)

## <a name="recommended-libraries-and-samples"></a>Рекомендуемые библиотеки и примеры

При доступе к Microsoft Graph без использования единого входа рекомендуется применять следующие библиотеки:

- Для надстроек, использующих серверную часть с платформой на основе .NET, например .NET Core или ASP.NET, используйте [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Для надстроек, использующих серверную часть на основе NodeJS, используйте [Azure AD Passport](https://github.com/AzureAD/passport-azure-ad).
- Для надстроек, использующих неявный поток, используйте [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Дополнительные сведения о рекомендуемых библиотеках для работы с платформой удостоверений Майкрософт (ранее — AAD версии 2.0) см. в статье [Библиотеки проверки подлинности платформы удостоверений Майкрософт](/azure/active-directory/develop/reference-v2-libraries).

Следующие примеры получают данные Microsoft Graph из надстройки Office:

- [Надстройка Office в Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Outlook в Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Office в Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
