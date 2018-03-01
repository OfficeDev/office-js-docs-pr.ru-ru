---
title: Использование Office UI Fabric JS в надстройках Office
description: ''
ms.date: 12/04/2017
---

# <a name="use-office-ui-fabric-js-in-office-add-ins"></a>Использование Office UI Fabric JS в надстройках Office

Office UI Fabric — это интерфейсная платформа JavaScript для создания дизайна для Office и Office 365. Если вы создаете надстройку только с помощью JavaScript, не используя Angular или React, советуем использовать Fabric JS для создания пользовательского интерфейса. Дополнительные сведения см. в статье [Office UI Fabric JS](https://dev.office.com/fabric-js).

Из этой статьи вы узнаете, как использовать Fabric JS.  

## <a name="add-the-fabric-cdn-references"></a>Добавление ссылок на CDN Fabric
Чтобы сослаться на Fabric из CDN, добавьте на страницу приведенный ниже HTML-код.

```html
<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
<script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>
```

## <a name="use-fabric-js-ux-components"></a>Использование компонентов дизайна Fabric JS

В Fabric JS предоставлены некоторые компоненты дизайна, например кнопки и флажки, которые можно использовать в надстройке. Ниже приведен список компонентов дизайна Fabric JS, рекомендованных для надстроек. Чтобы использовать один из компонентов Fabric в надстройке, откройте документацию Fabric по ссылке и следуйте инструкциям в разделе **Использование этого компонента**. 

- [Строка навигации](https://dev.office.com/fabric-js/Components/Breadcrumb/Breadcrumb.html)
- [Кнопка](https://dev.office.com/fabric-js/Components/Button/Button.html) (рекомендуем использовать в надстройке маленькие кнопки. Добавляйте к ним поля по 16 пикселей, чтобы область прикосновения на сенсорных устройствах составляла по крайней мере 40 пикселей.)
- [Флажок](https://dev.office.com/fabric-js/Components/CheckBox/CheckBox.html)
- [ChoiceFieldGroup](https://dev.office.com/fabric-js/Components/ChoiceFieldGroup/ChoiceFieldGroup.html)
- [Выбор даты](https://dev.office.com/fabric-js/Components/DatePicker/DatePicker.html) (Пример реализации элемента управления "Выбор даты" в надстройке см. в примере кода [Отслеживание продаж в Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).)
- [Раскрывающееся меню](https://dev.office.com/fabric-js/Components/Dropdown/Dropdown.html)
- [Подпись](https://dev.office.com/fabric-js/Components/Label/Label.html)
- [Ссылка](https://dev.office.com/fabric-js/Components/Link/Link.html)
- [Список](https://dev.office.com/fabric-js/Components/List/List.html) (Рекомендуем изменить стили компонента по умолчанию в CSS.)
- [MessageBanner](https://dev.office.com/fabric-js/Components/MessageBanner/MessageBanner.html)
- [MessageBar](https://dev.office.com/fabric-js/Components/MessageBar/MessageBar.html)
- [Наложение](https://dev.office.com/fabric-js/Components/Overlay/Overlay.html)
- [Панель](https://dev.office.com/fabric-js/Components/Panel/Panel.html)
- [Сводка](https://dev.office.com/fabric-js/Components/Pivot/Pivot.html)
- [ProgressIndicator](https://dev.office.com/fabric-js/Components/ProgressIndicator/ProgressIndicator.html)
- [Окна поиска](https://dev.office.com/fabric-js/Components/SearchBox/SearchBox.html)
- [Индикатор работы](https://dev.office.com/fabric-js/Components/Spinner/Spinner.html)
- [Таблица](https://dev.office.com/fabric-js/Components/Table/Table.html)
- [TextField](https://dev.office.com/fabric-js/Components/TextField/TextField.html)
- [Переключатель](https://dev.office.com/fabric-js/Components/Toggle/Toggle.html)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>Обновление надстройки для использования Fabric JS
Если вы используете предыдущую версию Office UI Fabric и хотите перейти на Fabric JS, ознакомьтесь с новыми компонентами, интегрируйте их в надстройку и проверьте их работу. При планировании обновления обратите внимание на следующие моменты:

- Компоненты проще инициализировать с помощью Fabric JS. Чтобы инициализировать компонент в предыдущих версиях Fabric, файл JavaScript компонента Fabric, в том числе ссылку `<Script>` на этот файл, необходимо включить в проект надстройки. В Fabric JS больше не нужно включать файл JavaScript компонента Fabric и связанную ссылку `<Script>`. Все, что нужно сделать, — это инициализировать компонент.   
- Некоторые компоненты теперь предоставляют функции, определяющие поведение компонента дизайна. Например, элемент управления "флажок" имеет функцию `toggle` для переключения между состояниями "флажок установлен" и "флажок снят". 
- Обновлены некоторые имена классов и стили значков.
- Наиболее заметное изменение — использование элемента `<label>` во многих компонентах. Элемент `<label>` определяет стиль компонента. Для использования элемента `<label>` может потребоваться обновить код дизайна. Например, изменение значения выбранного атрибута элемента `<input>` для флажка JS Fabric никак не влияет на флажок. Вместо этого можно воспользоваться функцией `check`, `unCheck` или `toggle`.   

## <a name="implementation"></a>Реализация
Если вы ищете подробный пример кода с Fabric JS, просмотрите следующий ресурс:

- [Отслеживание продаж в Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

## <a name="see-also"></a>См. также
Если вы ищете примеры кода или документацию по предыдущей версии Fabric, просмотрите следующие ресурсы:

- [Конструктивные шаблоны (используется Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Пример пользовательского интерфейса Fabric для надстройки Office (используется Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Использование Fabric 2.6.1 в надстройке Office](ui-elements/using-office-ui-fabric.md)
 

