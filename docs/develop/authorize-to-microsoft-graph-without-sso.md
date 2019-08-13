---
title: Авторизация в Microsoft Graph без единого входа
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 0bf79daa74542d36d90976dfd3f699591a8646a6
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302963"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Авторизация в Microsoft Graph без единого входа

Вы можете получить для надстройки разрешение на доступ к данным Microsoft Graph, получив маркер доступа к Graph из Azure Active Directory (AAD). Это выполняется с помощью потока кода авторизации или неявного потока, как и в любом другом веб-приложении с одним исключением: служба AAD не разрешает открывать свою страницу входа в элементе iframe. При работе с надстройкой Office в *Office в Интернете* область задач является элементом iframe. Это означает, что вам потребуется открыть экран входа в AAD в диалоговом окне, вызванном с помощью Dialog API для Office. Это повлияет на способ использования библиотек помощника проверки подлинности и авторизации. Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).

Сведения о программной проверке подлинности с помощью AAD см. в статье [Общие сведения о платформе удостоверений Майкрософт (версии 2.0)](/azure/active-directory/develop/v2-overview). В этом наборе документов содержится много руководств и инструкций, а также ссылки на соответствующие примеры. Дополнительное напоминание: вам может потребоваться изменить код из примеров, чтобы запустить его в диалоговом окне Office, учитывая, что диалоговое окно запускается в отдельном процессе из области задач.

После получения кодом маркера доступа к Microsoft Graph он передает маркер доступа из диалогового окна в область задач или сохраняет маркер в базе данных и уведомляет область задач о доступности маркера в базе данных. (Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).) Код в области задач запрашивает данные из Microsoft Graph и включает маркер в эти запросы. Дополнительные сведения о вызовах Microsoft Graph и SDK для Microsoft Graph см. в статье [Документация Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Рекомендуемые библиотеки и примеры

При доступе к Microsoft Graph без использования единого входа рекомендуется применять следующие библиотеки:

- Для надстроек, использующих серверную часть с платформой на основе .NET, например .NET Core или ASP.NET, используйте [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Для надстроек, использующих серверную часть на основе NodeJS, используйте [Azure AD Passport](https://github.com/AzureAD/passport-azure-ad).
- Для надстроек, использующих неявный поток, используйте [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Дополнительные сведения о рекомендуемых библиотеках для работы с платформой удостоверений Майкрософт (ранее — AAD версии 2.0) см. в статье [Библиотеки проверки подлинности платформы удостоверений Майкрософт](/azure/active-directory/develop/reference-v2-libraries.md).

Следующие примеры получают данные Microsoft Graph из надстройки Office:

- [ASP.NET Microsoft Graph надстройки Office](https://github.com/OfficeDev/office-add-in-microsoft-graph-aspnet)
- [ASP.NET Microsoft Graph надстройки Outlook](https://github.com/OfficeDev/outlook-add-in-microsoft-graph-aspnet)

