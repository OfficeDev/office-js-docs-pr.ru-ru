---
title: Примеры кода надстроек Office
description: Список примеров кода надстроек Office, которые помогают научиться создавать собственные надстройки.
ms.date: 11/18/2021
localization_priority: high
ms.openlocfilehash: 74346226a73554501cae31c29632d9ec0b595f6f
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/19/2022
ms.locfileid: "62073313"
---
# <a name="office-add-in-code-samples"></a>Примеры кода надстроек Office

Эти примеры кода помогают узнать, как использовать различные возможности при разработке надстроек Office.

## <a name="getting-started"></a>Начало работы

В следующих примерах показано, как создать простейшую надстройку Office, содержащую только манифест, веб-страницу HTML и логотип. Эти компоненты являются основными частями надстройки Office. Дополнительные сведения о том, как начать работу, см. в наших [кратких руководствах](../quickstarts/excel-quickstart-jquery.md) и [учебниках](/search/?terms=tutorial&scope=Office%20Add-ins).

* [Надстройка Excel "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world)
* [Надстройка Outlook "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)
* [Надстройка PowerPoint "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world)
* [Надстройка Word "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world)

## <a name="outlook"></a>Outlook

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Использование активации Outlook на основе событий для пометки внешних получателей (предварительная версия)](/samples/officedev/Office-Add-in-samples/outlook-add-in-tag-external-recipients) | Используйте активацию на основе событий для запуска надстройки Outlook при изменении получателей в процессе создания сообщения. Надстройка также использует API `appendOnSendAsync` для добавления заявления об отказе. |
| [Использование активации Outlook на основе событий для задания подписи](/samples/officedev/Office-Add-in-samples/outlook-add-in-set-signature/) | Используйте активацию на основе событий для запуска надстройки Outlook при создании нового сообщения или встречи. Надстройка может отвечать на события, даже если область задач не открыта. Она также использует API `setSignatureAsync`. |

## <a name="excel"></a>Excel

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Открытие в Teams](/samples/officedev/Office-Add-in-samples/office-excel-add-in-open-in-teams/) | Создайте новую электронную таблицу Excel в Microsoft Teams, содержащую определенные вами данные.|
| [Вставка внешнего файла Excel и его заполнение данными JSON](/samples/officedev/Office-Add-in-samples/excel-add-in-insert-external-file/)  | Вставьте существующий шаблон из внешнего файла Excel в открытую книгу Excel. Затем заполните шаблон данными веб-службы JSON. |
| [Создание настраиваемых контекстных вкладок на ленте](/samples/officedev/Office-Add-in-samples/office-add-in-contextual-tabs/) | Создайте настраиваемую контекстную вкладку на ленте в пользовательском интерфейсе Office. В примере создается таблица: если пользователь перемещает фокус внутри нее, отображается настраиваемая вкладка. Если пользователь перемещается за ее пределы, настраиваемая вкладка будет скрыта. |
| [Использование сочетаний клавиш для действий надстройки Office](/samples/officedev/Office-Add-in-samples/office-add-in-keyboard-shortcuts) | Настройте базовый проект надстройки Excel с использованием сочетаний клавиш. |
| [Пример пользовательской функции, использующей рабочий веб-процесс](/samples/officedev/Office-Add-in-samples/excel-custom-function-web-worker-pattern/) | Используйте рабочие веб-процессы в пользовательских функциях, чтобы предотвратить блокировку пользовательского интерфейса надстройки Office. |
| [Использование методов хранения для доступа к данным из надстройки Office в автономном режиме](/samples/officedev/Office-Add-in-samples/use-storage-techniques-to-access-data-from-an-office-add-in-when-offline/) | Реализуйте localStorage, чтобы включить ограниченную функциональность для надстройки Office, если соединение прервано. |
| [Шаблон пакетной обработки пользовательских функций](/samples/officedev/Office-Add-in-samples/excel-custom-function-batching-pattern/)| Объедините несколько вызовов в один, чтобы уменьшить количество сетевых вызовов к удаленной службе.|

## <a name="shared-javascript-runtime"></a>Общая среда выполнения JavaScript

| Имя                | Описание         |
|:--------------------|:--------------------|
[Совместный доступ к глобальным данным с общей средой выполнения](/samples/officedev/Office-Add-in-samples/office-add-in-shared-runtime-global-data/) | Настройте базовый проект, в котором используется общая среда выполнения, для запуска кода для кнопок ленты, области задач и пользовательских функций в единой среде выполнения браузера. |
| [Управление лентой и пользовательским интерфейсом области задач и запуск кода при открытии документа](/samples/officedev/Office-Add-in-samples/office-add-in-ribbon-task-pane-ui/) | Создайте контекстные кнопки ленты, которые включаются в зависимости от состояния вашей надстройки. |

## <a name="authentication-authorization-and-single-sign-on-sso"></a>Проверка подлинности, авторизация и единый вход

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Пример надстройки Outlook с единым входом](/samples/officedev/Office-Add-in-samples/outlook-add-in-sso-aspnet/) | Используйте функцию единого входа в Office, чтобы предоставить надстройке доступ к данным Microsoft Graph.|
| [Получение данных OneDrive с помощью Microsoft Graph и msal.js в надстройке Office](/samples/officedev/Office-Add-in-samples/office-add-in-auth-graph-react/) | Создайте надстройку Office как одностраничное приложение без серверной части, которое подключается к Microsoft Graph, и получите доступ к книгам, хранящимся в OneDrive для бизнеса, чтобы обновить электронную таблицу.  |
| [Проверка подлинности надстройки Office для Microsoft Graph](/samples/officedev/Office-Add-in-samples/office-add-in-auth-aspnet-graph/) | Узнайте, как создать надстройку Microsoft Office, которая подключается к Microsoft Graph, и получить доступ к книгам, хранящимся в OneDrive для бизнеса, чтобы обновить электронную таблицу. |
| [Проверка подлинности надстройки Outlook для Microsoft Graph](/samples/officedev/Office-Add-in-samples/outlook-add-in-auth-aspnet-graph/) | Создайте надстройку Outlook, которая подключается к Microsoft Graph, и получите доступ к книгам, хранящимся в OneDrive для бизнеса, чтобы создать новое сообщение электронной почты. |
| [Надстройка Office с единым входом на ASP.NET](/samples/officedev/Office-Add-in-samples/office-add-in-sso-aspnet/) | Используйте API `getAccessToken` в Office.js, чтобы предоставить надстройке доступ к данным Microsoft Graph. Этот пример создан на основе ASP.NET. |
| [Надстройка Office с единым входом на Node.js](/samples/officedev/Office-Add-in-samples/office-add-in-sso-nodejs/) | Используйте API `getAccessToken` в Office.js, чтобы предоставить надстройке доступ к данным Microsoft Graph. Этот пример создан на основе Node.js.|

## <a name="additional-samples"></a>Дополнительные примеры

| Имя                | Описание         |
|:--------------------|:--------------------|
|[Использование общей библиотеки для переноса надстройки набора средств Visual Studio для Office в веб-надстройку Office](/samples/officedev/Office-Add-in-samples/vsto-shared-library-excel/) |Предоставляет стратегию повторного использования кода при переходе с надстроек VSTO на надстройки Office. |
| [Интеграция функции Azure с пользовательской функцией Excel](/samples/officedev/Office-Add-in-samples/azure-function-with-excel-custom-function/) | Интегрируйте функции Azure с пользовательскими функциями для перемещения в облако или интегрируйте дополнительные службы. |
|[Примеры динамического кода DPI](/samples/officedev/Office-Add-in-samples/dynamic-dpi-code-samples/) |Коллекция примеров для обработки изменений DPI в надстройках COM, VSTO и Office. |

## <a name="next-steps"></a>Дальнейшие действия

Присоединитесь к программе для разработчиков Microsoft 365. Получите бесплатные инструменты, песочницу и другие ресурсы, чтобы создавать решения для платформы Microsoft 365.

- [Бесплатная песочница для разработчиков](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Получите бесплатную 90-дневную подписку Microsoft 365 E5 для разработчиков с возможностью продления.
- [Примеры пакетов данных](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Автоматически настройте песочницу, установив пользовательские данные и содержимое для использования при создании решений.
- [Доступ к экспертам](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Получите доступ к мероприятиям сообщества, чтобы обучаться у экспертов по Microsoft 365.
- [Персонализированные рекомендации](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Быстро находите ресурсы для разработчиков с помощью персонализированной панели мониторинга.
