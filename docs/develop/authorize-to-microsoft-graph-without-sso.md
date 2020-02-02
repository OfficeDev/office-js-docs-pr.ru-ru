---
title: Авторизация в Microsoft Graph без единого входа
description: Узнайте, как осуществлять авторизацию в Microsoft Graph без единого входа
ms.date: 01/29/2020
localization_priority: Priority
ms.openlocfilehash: a7dbd1fdd85852a82fcb00050283fece84a04e04
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/01/2020
ms.locfileid: "41649972"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Авторизация в Microsoft Graph без единого входа

Ваша надстройка может получить разрешение на доступ к данным Microsoft Graph, получив маркер доступа к Graph из Azure Active Directory (AAD). Используйте поток кода авторизации или неявный поток как в любом другом веб-приложении с одним исключением: служба AAD не разрешает открывать свою страницу входа в элементе iframe. При работе с надстройкой Office в *Office в Интернете* область задач является элементом iframe. Это означает, что вам потребуется открыть экран входа в AAD в диалоговом окне, вызванном с помощью Dialog API для Office. Это повлияет на способ использования библиотек помощника проверки подлинности и авторизации. Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).

Сведения о программной проверке подлинности с помощью AAD см. в статье [Общие сведения о платформе удостоверений Майкрософт (версии 2.0)](/azure/active-directory/develop/v2-overview), где вы найдете учебники и руководства, а также ссылки на соответствующие примеры. Опять же, вам может потребоваться изменить код из примеров, чтобы запустить его в диалоговом окне Office, учитывая, что диалоговое окно Office запускается в отдельном процессе из области задач.

После получения кодом маркера доступа к Graph он передает маркер доступа из диалогового окна в область задач или сохраняет маркер в базе данных и уведомляет область задач о доступности маркера. (Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).) Код в области задач запрашивает данные из Graph и включает маркер в эти запросы. Дополнительные сведения о вызовах Graph и Graph SDK см. в статье [Документация Microsoft Graph](/graph/).

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
