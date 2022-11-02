---
title: Авторизация в Microsoft Graph из надстройки Office
description: Узнайте, как авторизоваться в Microsoft Graph из надстройки Office.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 37dd4be3acb92dc7884972de923d94936fa870f4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810171"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>Авторизация в Microsoft Graph из надстройки Office

Ваша надстройка может получить авторизацию для данных Microsoft Graph, получив маркер доступа к Microsoft Graph из платформа удостоверений Майкрософт. Используйте поток кода авторизации или неявный поток так же, как и в других веб-приложениях, но с одним исключением: платформа удостоверений Майкрософт не позволяет открывать страницу входа в iframe. Когда надстройка Office запущена в *Office в Интернете*, область задач является iframe. Это означает, что вам потребуется открыть страницу входа в диалоговом окне с помощью API диалогового окна Office. Это повлияет на способ использования библиотек помощника проверки подлинности и авторизации. Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).

> [!NOTE]
> Если вы реализуете единый вход и планируете получить доступ к Microsoft Graph, см [. статью Авторизация в Microsoft Graph с помощью единого входа](authorize-to-microsoft-graph.md).

Сведения о программе проверки подлинности с помощью платформа удостоверений Майкрософт см. в [документации по платформа удостоверений Майкрософт](/azure/active-directory/develop). В этом наборе документации вы найдете руководства и руководства, а также ссылки на соответствующие примеры. Опять же, вам может потребоваться изменить код в примерах для запуска в диалоговом окне Office, чтобы учесть диалоговое окно Office, которое выполняется отдельно от области задач.

После того как код получит маркер доступа в Microsoft Graph, он передает маркер доступа из диалогового окна в область задач или сохраняет маркер в базе данных и сообщает области задач о доступности маркера. (Дополнительные сведения см [. в статье Проверка подлинности с помощью API диалогового окна Office](auth-with-office-dialog-api.md) .) Код в области задач запрашивает данные из Microsoft Graph и включает маркер в эти запросы. Дополнительные сведения о вызове Microsoft Graph и пакетов SDK для Microsoft Graph см. в [документации по Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Рекомендуемые библиотеки и примеры

При доступе к Microsoft Graph рекомендуется использовать следующие библиотеки.

- Для надстроек, использующих серверную часть с платформой на основе .NET, например .NET Core или ASP.NET, используйте [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Для надстроек, использующих серверную часть на основе NodeJS, используйте [Azure AD Passport](https://github.com/AzureAD/passport-azure-ad).
- Для надстроек, использующих неявный поток, используйте [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Дополнительные сведения о рекомендуемых библиотеках для работы с платформой удостоверений Майкрософт (ранее — AAD версии 2.0) см. в статье [Библиотеки проверки подлинности платформы удостоверений Майкрософт](/azure/active-directory/develop/reference-v2-libraries).

Следующие примеры получают данные Microsoft Graph из надстройки Office.

- [Надстройка Office в Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Outlook в Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Office в Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
