---
title: Office UI Fabric в надстройках Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b573f720ebe4f90f7d4dbfdb05693871b93a2258
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925194"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric в надстройках Office 

Office UI Fabric — это интерфейсная платформа JavaScript для создания дизайна для Office и Office 365. В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office. Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно. 

Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.

В следующих разделах описано, как начать использовать Fabric для своих потребностей. 

## <a name="use-fabric-core-icons-fonts-colors"></a>Использование Fabric Core: значки, шрифты, цвета
Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку. Fabric Core не зависит от платформы. И Fabric React, и Fabric JS используют Fabric Core.

Чтобы начать работу с Fabric Core:

1. Добавьте ссылку CDN в HTML-код на своей странице.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. Используйте значки и шрифты Fabric. 

    Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7). 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://dev.office.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`. 

    Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://dev.office.com/fabric#/styles/typography) и [Цвета](https://dev.office.com/fabric#/styles/colors).
 
## <a name="use-fabric-components"></a>Использование компонентов Fabric 
В Fabric есть различные компоненты оформления, которые можно использовать при создании надстроек, в том числе:

- Компоненты ввода — например, Button, Checkbox и Toggle.
- Компоненты навигации — например, сводка и строка навигации
- Компоненты уведомления — например, MessageBar и Callout.  

Не все компоненты Fabric рекомендуются для использования в надстройках. Ниже приведен список компонентов Fabric React UX, которые мы рекомендуем использовать в надстройке:

- [Строка навигации](https://developer.microsoft.com/fabric#/components/breadcrumb)
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
|**Angular**| См. проект сообщества [ngOfficeUIFabric](http://ngofficeuifabric.com/) с директивами Angular 1.5 и раздел [Обдумайте размещение компонентов Fabric в компонентах Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
