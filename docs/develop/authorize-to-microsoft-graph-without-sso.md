---
title: Авторизуйте Graph Майкрософт из Office надстройки
description: Узнайте, как авторизировать Graph майкрософт из Office надстройки.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8166b7a71767abd0456662dbe8573f59bb2c7e82
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743582"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>Авторизуйте Graph Майкрософт из Office надстройки

Ваша надстройка может получить авторизацию в Microsoft Graph данных, получив маркер доступа к microsoft Graph из платформа удостоверений Майкрософт. Используйте поток кода авторизации или неявный поток так же, как и в других веб-приложениях, но за одним исключением: платформа удостоверений Майкрософт не позволяет открыть свою страницу регистрации в iframe. При работе с надстройкой Office в *Office в Интернете* область задач является элементом iframe. Это означает, что вам потребуется открыть страницу входной двери в диалоговом окне с помощью API Office диалогов. Это повлияет на способ использования библиотек помощника проверки подлинности и авторизации. Дополнительные сведения см. в статье [Проверка подлинности с помощью Dialog API для Office](auth-with-office-dialog-api.md).

> [!NOTE]
> Если вы реализуете SSO и планируете доступ к microsoft Graph, см. в Graph с [SSO](authorize-to-microsoft-graph.md).

Сведения о проверке подлинности с помощью платформа удостоверений Майкрософт см. в платформа удостоверений Майкрософт [документации](/azure/active-directory/develop). В этом наборе документации вы найдете учебники и руководства, а также ссылки на соответствующие примеры. Еще раз, возможно, потребуется настроить код в примерах для запуска в диалоговом окне Office, чтобы учитывать диалоговое окно Office, которое выполняется в отдельном процессе от области задач.

После получения маркера доступа в Microsoft Graph код передает маркер доступа из диалогового окна в области задач, либо сохраняет маркер в базе данных и сигнализирует области задач о том, что маркер доступен. ([Подробные сведения см. в Office диалоговом API](auth-with-office-dialog-api.md).) Код в области задач запрашивает данные из Microsoft Graph и включает маркер в эти запросы. Дополнительные сведения о вызове microsoft Graph и SDKs microsoft Graph см. в документации [microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Рекомендуемые библиотеки и примеры

Рекомендуется использовать следующие библиотеки при доступе к microsoft Graph.

- Для надстроек, использующих серверную часть с платформой на основе .NET, например .NET Core или ASP.NET, используйте [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Для надстроек, использующих серверную часть на основе NodeJS, используйте [Azure AD Passport](https://github.com/AzureAD/passport-azure-ad).
- Для надстроек, использующих неявный поток, используйте [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Дополнительные сведения о рекомендуемых библиотеках для работы с платформой удостоверений Майкрософт (ранее — AAD версии 2.0) см. в статье [Библиотеки проверки подлинности платформы удостоверений Майкрософт](/azure/active-directory/develop/reference-v2-libraries).

В следующих примерах Graph microsoft Office надстройки.

- [Надстройка Office в Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Outlook в Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Office в Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
