# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric в надстройках Office 

Office UI Fabric — это интерфейсная платформа JavaScript для создания дизайна для Office и Office 365. В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office. Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно. 

Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.

В следующих разделах описано, как начать использовать Fabric для своих потребностей. 

## <a name="use-fabric-core-icons-fonts-colors"></a>Использование Fabric Core: значки, шрифты, цвета
Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку. Fabric Core не зависит от платформы. И Fabric React, и Fabric JS используют Fabric Core.

Чтобы начать работу с Fabric Core:

1. Добавьте ссылку CDN в HTML-код на своей странице.  

    `<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">`   
    
2. Используйте значки и шрифты Fabric. 

Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7). 
   
`<i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>`

Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://dev.office.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`. 

Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://dev.office.com/fabric#/styles/typography) и [Цвета](https://dev.office.com/fabric#/styles/colors).
 
## <a name="use-fabric-components"></a>Использование компонентов Fabric 
В Fabric есть различные компоненты оформления, которые можно использовать при создании надстроек, в том числе:

- Компоненты ввода — например, Button, Checkbox и Toggle.
- Компоненты навигации — например, Pivot Breadcrumb.
- Компоненты уведомления — например, MessageBar и Callout.  

Не все компоненты Fabric рекомендуется использовать в надстройках. Этот раздел содержит инструкции по использованию рекомендуемых компонентов. Например, инструкции по использованию кнопки Fabric в надстройке приведены в разделе [Кнопка](button.md). 

Для создания надстройки можно использовать разные платформы JavaScript, такие как Angular или React. Прежде чем использовать компоненты Fabric со своей платформой, ознакомьтесь с перечисленными ниже ресурсами.

|**Платформа**|**Пример**|
|:------------|:----------|
|**Только JavaScript** (без платформы)|[Использование Office UI Fabric JS в надстройках Office](using-office-ui-fabric-js.md).|
|**React**|[Использование Office UI Fabric React в надстройках Office](using-office-ui-fabric-react.md )|
|**Angular**| См. проект сообщества [ngOfficeUIFabric](http://ngofficeuifabric.com/) с директивами Angular 1.5 и раздел [Обдумайте размещение компонентов Fabric в компонентах Angular 2](https://dev.office.com/docs/add-ins/develop/add-ins-with-angular2#consider-wrapping-fabric-components-with-angular-2-components)|