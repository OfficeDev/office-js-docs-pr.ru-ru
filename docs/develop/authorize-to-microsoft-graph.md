---
title: Авторизованный доступ в Microsoft Graph из вашей надстройки Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: f6e7de146d2f03256aa673a0653c1e03f9340d86
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876594"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Авторизованный доступ в Microsoft Graph из вашей надстройки Office (предварительная версия)

Пользователи входят в Office (в Интернете, на мобильных устройствах и настольных компьютерах), используя личную учетную запись Майкрософт либо рабочую или учебную учетную запись (Office 365). Лучший способ предоставления авторизованного доступа в [Microsof Graph](https://developer.microsoft.com/graph/docs) из надстройки Office — использование учетных данных для входа непосредственно в Office. В результате пользователь получает доступ к своим данным Microsoft Graph без необходимости повторного входа в систему. 

> [!NOTE]
> В настоящее время API единого входа поддерживается в тестовом режиме для Word, Excel, Outlook и PowerPoint. Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API идентификации](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).
> Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Архитектура надстройки для единого входа и Microsoft Graph

Помимо страниц и кода JavaScript веб-приложения, в надстройке также должны размещаться (с тем же [полным доменном именем](https://msdn.microsoft.com/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly)) один или несколько веб-API, которые будут получать маркер доступа и отправлять запросы к Microsoft Graph.

Манифест надстройки содержит разметку, которая указывает, как надстройка регистрируется в конечной точке Azure Active Directory (Azure AD) версии 2.0, а также задает необходимые надстройке разрешения для Microsoft Graph.

### <a name="how-it-works-at-runtime"></a>Принцип работы в среде выполнения

На следующей схеме изображен процесс входа в систему и получения доступа к Microsoft Graph.

![Схема единого входа](../images/sso-access-to-microsoft-graph.png)

1. Код JavaScript надстройки вызывает новый API Office.js [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Он указывает ведущему приложению Office, что необходимо получить маркер доступа к надстройке (далее он будет называться **маркером доступа к начальной загрузке**, поскольку в этом процессе он позже будет заменен вторым маркером. Пример декодированного маркера доступа к начальной загрузке см. в разделе [Пример маркера доступа](sso-in-office-add-ins.md#example-access-token)).
1. Если вход в Office не выполнен, в ведущем приложении открывается всплывающее окно, в котором пользователю предлагается войти.
1. Если текущий пользователь запускает надстройку в первый раз, ему предлагается дать согласие.
1. Ведущее приложение Office запрашивает **маркер доступа к начальной загрузке** у конечной точки Azure AD версии 2.0 для текущего пользователя.
1. Azure AD отправляет маркер начальной загрузки ведущему приложению Office.
1. Ведущее приложение Office отправляет **маркер доступа к начальной загрузке** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessTokenAsync`.
1. Код JavaScript надстройки отправляет HTTP-запрос к веб-API с тем же полным доменным именем, что и у надстройки. Этот запрос включает **маркер доступа к начальной загузке** в качестве подтверждения авторизации.  
1. Серверный код проверяет входящий **маркер доступа к начальной загрузке**.
1. Серверный код использует поток "от имени»" (определенный в [Обмен токенами OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) и [демон или серверное приложение для сценария веб-API Azure](https://docs.microsoft.com/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)), чтобы получить маркер доступа для Microsoft Graph в обмен на маркер доступа к начальной загрузке.
1. Azure AD возвращает в надстройку маркер доступа для Microsoft Graph (и маркер обновления, если надстройка запрашивает разрешение *offline_access*).
1. Серверный код кэширует маркер доступа в Microsoft Graph.
1. Серверный код отправляет запросы в Microsoft Graph и включает маркер доступа в Microsoft Graph.
1. Microsoft Graph возвращает данные надстройке, а она может передать их в свой пользовательский интерфейс.
1. Когда маркер доступа в Microsoft Graph истекает, серверный код может использовать свой маркер обновления, чтобы получить новый маркер доступа в Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Разработка надстройки единого входа для Microsoft Graph

Надстройка, выполняющая вход Microsoft Graph, разрабатывается так же, как и любая другая надстройка с единым входом. Подробное описание см. в разделе [Включение единого входа для надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins). Разница заключается в том, что надстройка должна обязательно иметь веб-API на стороне сервера, а также в использовании термина "маркер доступа к начальной загрузке" вместо термина "маркер доступа" из указанной статьи. 

В зависимости от языка и платформы могут быть доступны библиотеки, которые упростят создание нужного серверного кода. Ваш код должен выполнять следующие операции:

* Проверять в настройке маркер доступа к начальной загрузке, полученный от созданного ранее обработчика маркеров. Подробнее см. в статье [Проверка маркера доступа](sso-in-office-add-ins.md#validate-the-access-token). 
* Запускать поток "от имени" путем вызова конечной точки Azure AD версии 2.0, в ходе которого передается маркер доступа к надстройке, некоторые метаданные пользователя и учетные данные надстройки (идентификатор и секрет).
* Помещать в кэш возвращенный маркер доступа для Microsoft Graph. Подробнее об этом потоке см. в статье [Azure Active Directory v2.0 и поток "от имени" OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Создавать один или несколько методов веб-API, которые получают данные Microsoft Graph, передавая кэшированный маркер доступа в Microsoft Graph.

> [!NOTE]
> Примеры декодированных маркеров доступа для Microsoft Graph, которые были получены потоком "от имени", см. в статье [Azure Active Directory v2.0 и поток "от имени" OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Примеры подробных пошаговых инструкций и сценариев см. ниже:

* [Создание надстройки Office на платформе Node.js с использованием единого входа](create-sso-office-add-ins-nodejs.md)
* [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md)
* [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



