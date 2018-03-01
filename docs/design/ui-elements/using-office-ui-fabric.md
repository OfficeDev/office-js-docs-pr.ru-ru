---
title: Использование Office UI Fabric 2.6.1 в надстройках Office
description: ''
ms.date: 12/04/2017
---



# <a name="use-office-ui-fabric-261-in-office-add-ins"></a>Использование Office UI Fabric 2.6.1 в надстройках Office

При создании надстроек Office мы рекомендуем использовать [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) для разработки пользовательского интерфейса. Далее представлены основные принципы использования Office UI Fabric.  

> [!NOTE]
> Информацию об Office UI Fabric JS см. в статье [Использование Office UI Fabric в надстройках Office](../using-office-ui-fabric-js.md).

## <a name="1-set-up-fabric"></a>1. Настройка Fabric

Добавьте следующие строки в код HTML в разделе head, чтобы указать ссылку на Office UI Fabric из сети CDN.

```HTML
<link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
<link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">
```


## <a name="2-use-fabric-icons-and-fonts"></a>2. Использование значков и шрифтов Office UI Fabric

Использовать значки очень просто. Все что нужно — указать элемент "i" и добавить ссылку на соответствующие классы. Вы можете задать размер значка, изменив размер шрифта.

```HTML
<i class="ms-Icon ms-Icon--group" style="font-size:xx-large" aria-hidden="true"></i>
```


## <a name="3-use-styles-for-simple-components"></a>3. Использование стилей для простых компонентов

Office UI Fabric поставляется со стилями для различных элементов пользовательского интерфейса, таких как кнопки и флажки. Все что вам нужно — указать ссылку на соответствующие классы, чтобы добавить нужный стиль, как показано в приведенном ниже примере.

```HTML
<button class="ms-Button" id="get-data-from-selection">
<span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
<span class="ms-Button-label">Get Data from selection</span>
<span class="ms-Button-description">Get Data from the document selection</span>
</button>
```

## <a name="4-use-components-with-sample-behavior"></a>4. Использование компонентов с примерами поведения

Office UI Fabric включает некоторые компоненты, поддерживающие определенное поведение (например, необходимое в случае щелчка). Мы добавили в **Fabric 2.6.1** **примеры кода** в форме подключаемых модулей интерфейса JQuery. Воспользуйтесь ими, чтобы начать работу. Кроме того, можно применять и любую другую платформу. Если вы решили использовать образцы, обратите внимание, что они не распространяются по сети CDN, поэтому вам необходимо скачать их в [разделе сайта GitHub, посвященном Fabric](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1) (**выпуск 2.6.1**), указать на них ссылку, а затем реализовать в своем коде. 

Например, чтобы использовать компонент SearchBox:

1. Скачайте компонент SearchBox с [GitHub](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1/src/components/SearchBox).
2. Добавьте следующую ссылку в код: `<script src="SearchBox/Jquery.SearchBox.js"></script>`.
3. Инициализируйте компонент, убедившись, что эта строка выполняется при загрузке страницы: `$(".ms-SearchBox").SearchBox();`. Рекомендуется включить этот код в блок `Office.Initialize` вашей надстройки.     

> [!NOTE]
> Если вы не собираетесь использовать все компоненты Fabric, вы можете скачивать меньше ресурсов, разместив отдельные CSS-файлы для каждого компонента. Вы можете скопировать CSS-файлы из папок компонентов в [репозитории GitHub для Fabric 2.6.1](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1). 


## <a name="next-steps"></a>Дальнейшие действия

Ищете подробные примеры использования Fabric? У нас есть, что показать. Ознакомьтесь с [примером пользовательского интерфейса Fabric для надстройки Office](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample). Вы также можете посетить интерактивный веб-сайт [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric).

## <a name="see-also"></a>См. также

- [Рекомендации по проектированию надстроек Office](../add-in-design.md)