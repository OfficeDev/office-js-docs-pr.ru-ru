---
title: Примеры кода надстроек Office
description: Список примеров кода надстроек Office, которые помогают научиться создавать собственные надстройки.
ms.date: 06/10/2022
localization_priority: high
ms.openlocfilehash: 16a1f92992c397772559468c27033aa58f6b6a6d
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423267"
---
# <a name="office-add-in-code-samples"></a>Примеры кода надстроек Office

Эти примеры кода помогают узнать, как использовать различные возможности при разработке надстроек Office.

## <a name="getting-started"></a>Начало работы

В следующих примерах показано, как создать простейшую надстройку Office, содержащую только манифест, веб-страницу HTML и логотип. Эти компоненты являются основными частями надстройки Office. Дополнительные сведения о том, как начать работу, см. в наших [кратких руководствах](../quickstarts/excel-quickstart-jquery.md) и [учебниках](/search/?terms=tutorial&scope=Office%20Add-ins).

- [Надстройка Excel "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world)
- [Надстройка Outlook "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)
- [Надстройка PowerPoint "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world)
- [Надстройка Word "Hello world"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world)

<br>

---

---

## <a name="blazor-webassembly"></a>Blazor WebAssembly

- [Создание надстройки Blazor WebAssembly для Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/excel-blazor-add-in)
- [Создание надстройки Blazor WebAssembly для Word](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/word-blazor-add-in)

## <a name="excel"></a>Excel

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Открытие в Teams](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-open-in-teams) | Создайте новую электронную таблицу Excel в Microsoft Teams, содержащую определенные вами данные.|
| [Вставка внешнего файла Excel и его заполнение данными JSON](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-insert-file)  | Вставьте существующий шаблон из внешнего файла Excel в открытую книгу Excel. Затем заполните шаблон данными веб-службы JSON. |
| [Создание настраиваемых контекстных вкладок на ленте](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs) | Создайте настраиваемую контекстную вкладку на ленте в пользовательском интерфейсе Office. В примере создается таблица: если пользователь перемещает фокус внутри нее, отображается настраиваемая вкладка. Если пользователь перемещается за ее пределы, настраиваемая вкладка будет скрыта. |
| [Использование сочетаний клавиш для действий надстройки Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) | Настройте базовый проект надстройки Excel с использованием сочетаний клавиш. |
| [Пример пользовательской функции, использующей рабочий веб-процесс](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/web-worker) | Используйте рабочие веб-процессы в пользовательских функциях, чтобы предотвратить блокировку пользовательского интерфейса надстройки Office. |
| [Использование методов хранения для доступа к данным из надстройки Office в автономном режиме](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin) | Реализуйте localStorage, чтобы включить ограниченную функциональность для надстройки Office, если соединение прервано. |
| [Шаблон пакетной обработки пользовательских функций](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching)| Объедините несколько вызовов в один, чтобы уменьшить количество сетевых вызовов к удаленной службе.|

## <a name="outlook"></a>Outlook

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Шифрование вложений, обработка участников в приглашениях на собрания и реагирование на изменения даты и времени встречи](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-encrypt-attachments) | Используйте активацию на основе событий для шифрования вложений при добавлении пользователем. Также используйте обработку событий для получателей, измененных в приглашении на собрание, и изменений даты и времени начала и окончания в приглашении на собрание. |
| [Использование активации Outlook на основе событий для пометки внешних получателей](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external) | Используйте активацию на основе событий для запуска надстройки Outlook при изменении получателей в процессе создания сообщения. Надстройка также использует API `appendOnSendAsync` для добавления заявления об отказе. |
| [Использование активации Outlook на основе событий для задания подписи](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature) | Используйте активацию на основе событий для запуска надстройки Outlook при создании нового сообщения или встречи. Надстройка может отвечать на события, даже если область задач не открыта. Она также использует API `setSignatureAsync`. |
| [Использование интеллектуальных оповещений Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories) | Используйте интеллектуальные оповещения Outlook, чтобы убедиться, что требуемые цветовые категории применяются к новому сообщению или встрече перед его отправкой. |

## <a name="word"></a>Word

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Получение, редактирование и настройка OOXML-содержимого в документе Word с помощью надстройки Word](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-get-set-edit-openxml) | В этом примере показано, как получать, редактировать и настраивать OOXML-содержимое в документе Word. Пример надстройки предоставляет электронный блокнот для получения собственного контента в формате Office Open XML, а также тестирования собственных фрагментов Office Open XML.|
| [Загрузка и запись содержимого в формате Open XML в надстройке Word](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)  | В этом примере надстройки показано, как добавлять форматированное содержимое различных типов в документ Word с помощью метода setSelectedDataAsync с типом приведения ooxml. С помощью этой надстройки также можно показывать разметку Office Open XML для каждого типа контента в примере прямо на странице. |

<br>

---

---

## <a name="authentication-authorization-and-single-sign-on-sso"></a>Проверка подлинности, авторизация и единый вход

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Пример надстройки Outlook с единым входом](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) | Используйте функцию единого входа в Office, чтобы предоставить надстройке доступ к данным Microsoft Graph.|
| [Получение данных OneDrive с помощью Microsoft Graph и msal.js в надстройке Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React) | Создайте надстройку Office как одностраничное приложение без серверной части, которое подключается к Microsoft Graph, и получите доступ к книгам, хранящимся в OneDrive для бизнеса, чтобы обновить электронную таблицу.  |
| [Проверка подлинности надстройки Office для Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) | Узнайте, как создать надстройку Microsoft Office, которая подключается к Microsoft Graph, и получить доступ к книгам, хранящимся в OneDrive для бизнеса, чтобы обновить электронную таблицу. |
| [Проверка подлинности надстройки Outlook для Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET) | Создайте надстройку Outlook, которая подключается к Microsoft Graph, и получите доступ к книгам, хранящимся в OneDrive для бизнеса, чтобы создать новое сообщение электронной почты. |
| [Надстройка Office с единым входом на ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) | Используйте API `getAccessToken` в Office.js, чтобы предоставить надстройке доступ к данным Microsoft Graph. Этот пример создан на основе ASP.NET. |
| [Надстройка Office с единым входом на Node.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | Используйте API `getAccessToken` в Office.js, чтобы предоставить надстройке доступ к данным Microsoft Graph. Этот пример создан на основе Node.js.|

## <a name="shared-runtime"></a>Общее время выполнения

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Совместный доступ к глобальным данным с общей средой выполнения](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-global-state) | Настройте базовый проект, в котором используется общая среда выполнения, для запуска кода для кнопок ленты, области задач и пользовательских функций в единой среде выполнения браузера. |
| [Управление лентой и пользовательским интерфейсом области задач и запуск кода при открытии документа](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario) | Создайте контекстные кнопки ленты, которые включаются в зависимости от состояния вашей надстройки. |

<br>

---

---

## <a name="additional-samples"></a>Дополнительные примеры

| Имя                | Описание         |
|:--------------------|:--------------------|
| [Использование общей библиотеки для переноса надстройки набора средств Visual Studio для Office в веб-надстройку Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/VSTO-shared-code-migration) | Предоставляет стратегию повторного использования кода при переходе с надстроек VSTO на надстройки Office. |
| [Интеграция функции Azure с пользовательской функцией Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AzureFunction) | Интегрируйте функции Azure с пользовательскими функциями для перемещения в облако или интегрируйте дополнительные службы. |
| [Примеры динамического кода DPI](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/dynamic-dpi) | Коллекция примеров для обработки изменений DPI в надстройках COM, VSTO и Office. |

## <a name="next-steps"></a>Дальнейшие действия

Присоединитесь к программе для разработчиков Microsoft 365. Получите бесплатные инструменты, песочницу и другие ресурсы, чтобы создавать решения для платформы Microsoft 365.

- [Бесплатная песочница для разработчиков](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Получите бесплатную 90-дневную подписку Microsoft 365 E5 для разработчиков с возможностью продления.
- [Примеры пакетов данных](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Автоматически настройте песочницу, установив пользовательские данные и содержимое для использования при создании решений.
- [Доступ к экспертам](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Получите доступ к мероприятиям сообщества, чтобы обучаться у экспертов по Microsoft 365.
- [Персонализированные рекомендации](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Быстро находите ресурсы для разработчиков с помощью персонализированной панели мониторинга.
