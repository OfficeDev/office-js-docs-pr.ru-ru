---
title: Fabric Core в надстройках Office
description: Обзор использования компонентов Fabric Core и пользовательского интерфейса Fabric в Office надстройки.
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 3d10cc5d8f33c8dd66f4f988fdd5a082580b1aca
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074205"
---
# <a name="fabric-core-in-office-add-ins"></a>Fabric Core в надстройках Office

Fabric Core — это коллекция классов CSS и mixins SASS, предназначенных для использования в надстройки React *Office.* Fabric Core содержит основные элементы языка Fluent пользовательского интерфейса, такие как значки, цвета, шрифты и сетки. Fabric Core является независимой структурой, поэтому ее можно использовать с любым одно-страницным приложением или любой серверной веб-структурой пользовательского интерфейса. (Он называется "Fabric Core" вместо "Fluent Core" по историческим причинам.)

Если пользовательский интерфейс надстройки не React, вы также можете использовать набор компонентов, не React. См. [Office UI Fabric компоненты JS.](#use-office-ui-fabric-js-components)

> [!NOTE]
> В этой статье описывается использование Fabric Core в контексте Office надстройки. Но он также используется в широком диапазоне Microsoft 365 приложений и расширений. Дополнительные сведения см. в [материале Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) и репо [с открытым исходным кодом Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).

## <a name="use-fabric-core-icons-fonts-colors"></a>Использование Fabric Core: значки, шрифты, цвета

1. Добавьте ссылку сети доставки контента (CDN) на HTML на своей странице.

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Используйте значки и шрифты Fabric Core.

    Чтобы использовать значок Fabric Core, включите элемент "i" на странице, а затем со ссылкой на соответствующие классы. Вы можете изменять размер значка, изменяя размер шрифта. Например, ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Дополнительные инструкции см. [в Fluent UI Icons.](https://developer.microsoft.com/fluentui#/styles/web/icons) Чтобы найти дополнительные значки, доступные в Fabric Core, используйте функцию поиска на этой странице. Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.

    Сведения о размерах шрифтов и цветах, доступных в Fabric Core, см. в [typography](https://developer.microsoft.com/fluentui#/styles/web/typography) и таблице **Цветов** содержимого [в Цветах.](https://developer.microsoft.com/fluentui#/styles/web/colors)

Примеры, включенные в [примеры, далее](#samples) в этой статье.

## <a name="use-office-ui-fabric-js-components"></a>Использование Office UI Fabric JS

Надстройки с React UIS также могут использовать все многие компоненты [из Office UI Fabric JS,](https://github.com/OfficeDev/office-ui-fabric-js)включая кнопки, диалоги, выборщики и многое другое. Ознакомьтесь с чтением репо для инструкций.

Примеры, включенные в [примеры, далее](#samples) в этой статье.

## <a name="samples"></a>Примеры

В следующем примере надстройки используют компоненты Fabric Core Office UI Fabric JS. Некоторые из этих репозитов архивироваться, что означает, что они больше не обновляются с помощью ошибок или исправлений безопасности, но вы все еще можете использовать их, чтобы узнать, как использовать компоненты пользовательского интерфейса Fabric Core и Fabric.

- [Excel Надстройка JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excel надстройки SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excel тенденции расходов на надстройку WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excel надстройка контента Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Office пример пользовательского интерфейса Надстройки Fabric](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlook Надстройка GifMe](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPoint надстройки Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
