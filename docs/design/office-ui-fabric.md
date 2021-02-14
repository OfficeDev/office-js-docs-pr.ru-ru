---
title: Office UI Fabric в надстройках Office
description: Обзор использования компонентов Office UI Fabric в надстройки Office.
ms.date: 2/09/2021
localization_priority: Normal
ms.openlocfilehash: 9799d98d795486203e4bcc23bffc043c2ead6e28
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237681"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric в надстройках Office

Office UI Fabric — это интерфейсная структура JavaScript для создания пользовательского интерфейса для Office. В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office. Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно.

Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.

В следующих разделах описано, как начать использовать Fabric для своих потребностей.

## <a name="use-fabric-core-icons-fonts-colors"></a>Использование Fabric Core: значки, шрифты, цвета

Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку. Fabric Core не зависит от платформы. Fabric Core используется с помощью Fabric React и входит в его состав.

Чтобы начать работу с Fabric Core:

1. Добавьте ссылку CDN в HTML-код на своей странице.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Используйте значки и шрифты Fabric.

    Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://developer.microsoft.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.

    Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://developer.microsoft.com/fabric#/styles/typography) и [Цвета](https://developer.microsoft.com/fabric#/styles/colors).

## <a name="use-fabric-components"></a>Использование компонентов Fabric

Fabric предоставляет различные компоненты UX, которые можно использовать для создания надстройки. Мы не ожидаем, что все компоненты Fabric будут использоваться одной надстройой. Определите оптимальные компоненты для сценария и пользовательского интерфейса (например, может быть сложно правильно отобразить навигацию в области задач). [](https://developer.microsoft.com/fabric#/components/breadcrumb)

Ниже приводится список распространенных компонентов [fabric React UX,](https://developer.microsoft.com/fluentui#/controls/web) которые мы рекомендуем использовать в надстройки.

- [Кнопка](https://developer.microsoft.com/fabric#/components/button)
- [Флажок](https://developer.microsoft.com/fabric#/components/checkbox)
- [ChoiceGroup](https://developer.microsoft.com/fabric#/components/choicegroup)
- [Раскрывающееся меню](https://developer.microsoft.com/fabric#/components/dropdown)
- [Подпись](https://developer.microsoft.com/fabric#/components/label)
- [Список](https://developer.microsoft.com/fabric#/components/list)
- [Сводка](https://developer.microsoft.com/fabric#/components/pivot)
- [TextField](https://developer.microsoft.com/fabric#/components/textfield)
- [Переключатель](https://developer.microsoft.com/fabric#/components/toggle)

Для создания надстройки можно использовать разные платформы JavaScript, такие как Angular или React. Прежде чем использовать компоненты Fabric со своей платформой, ознакомьтесь с перечисленными ниже ресурсами.

|**Платформа**|**Пример**|
|:------------|:----------|
|**React**|[Использование Office UI Fabric React в надстройках Office](using-office-ui-fabric-react.md )|
|**Angular**| [Рассмотрите возможность переноса компонентов Fabric с помощью компонентов Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
