---
title: Запрос разрешений на использование API в контентных надстройках и надстройках области задач
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: eb80c0b18848da9f0844ae3eef5f3c5dc467d932
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457966"
---
# <a name="requesting-permissions-for-api-use-in-content-and-task-pane-add-ins"></a>Запрос разрешений на использование API в контентных надстройках и надстройках области задач

В этой статье описываются различные уровни разрешений, которые можно объявить в манифесте контентной надстройки или надстройки области задач, чтобы указать уровень доступа JavaScript API, необходимый вашей надстройке. 




## <a name="permissions-model"></a>Модель разрешений


5-уровневая модель разрешений JavaScript API — это основа безопасности и конфиденциальности для пользователей контентных надстроек и надстроек области задач. На рис. 1 показаны пять уровней разрешений API, которые можно объявить в манифесте надстройки.


*Рис. 1. Пятиуровневая модель разрешений для контентных надстроек и надстроек области задач*

![Уровни разрешений для приложений области задач](../images/office15-app-sdk-task-pane-app-permission.png)



Эти разрешения задают подмножество API, которые контентная надстройки или надстройка области задач сможет использовать во время выполнения, когда пользователь вставляет, а затем активирует ваше приложение (доверяет ему). Чтобы объявить уровень разрешений, необходимый вашей надстройке, укажите одно из текстовых значений разрешения в элементе [Permissions](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions) манифеста надстройки. Следующий пример запрашивает разрешение **WriteDocument**, который разрешает использовать только методы записи в документ (но не методы чтения).




```XML
<Permissions>WriteDocument</Permissions>
```

Рекомендуется запрашивать разрешения по принципу  _минимальной привилегии_. Т. е. следует запрашивать разрешения для доступа только к минимальному подмножеству API, необходимому надстройке для правильной работы. Например, если надстройке требуется только читать данные документа пользователя, оно должно запрашивать максимум разрешение **ReadDocument**.

В следующей таблице описывается подмножество API JavaScript, предоставляемое каждым уровнем разрешений.



|**Разрешение**|**Включенное подмножество API**|
|:-----|:-----|
|**Существуют ограничения**|Методы объекта [Settings](https://docs.microsoft.com/javascript/api/office/office.settings) и метод [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document#getactiveviewasync-options--callback-).Это минимальный уровень разрешений, запрашиваемый контентной надстройкой или надстройкой области задач.|
|**ReadDocument**|В дополнение к интерфейсам API, предоставляемым разрешением  **Restricted**, добавляет доступ к элементам API, необходимым для чтения документа и управления привязками.В том числе разрешается использование следующих элементов:<br/><ul><li>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-" target="_blank">Document.getSelectedDataAsync</a> для получения выбранного текста, HTML (только Word) или табличных данных, но не базовый код Open Office XML (OOXML), содержащий все данные в документе.</p></li><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.document#getfileasync-filetype--options--callback-" target="_blank">Document.getFileAsync</a> для получения всего текста документа, но не двоичной копии OOXML документа.</p></li><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.binding#getdataasync-options--callback-" target="_blank">Binding.getDataAsync</a> для чтения связанных данных в документе.</p></li><li><p>Методы <a href="https://docs.microsoft.com/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-" target="_blank">addFromNamedItemAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-" target="_blank">addFromPromptAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-" target="_blank">addFromSelectionAsync</a> объекта <span class="keyword">Bindings</span> для создания привязок в документе.</p></li><li><p>Методы <a href="https://docs.microsoft.com/javascript/api/office/office.bindings#getallasync-options--callback-" target="_blank">getAllAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.bindings#getbyidasync-id--options--callback-" target="_blank">getByIdAsync</a> и <a href="https://docs.microsoft.com/javascript/api/office/office.bindings#releasebyidasync-id--options--callback-" target="_blank">releaseByIdAsync</a> объекта <span class="keyword">Bindings</span> для доступа к привязкам документа и их удаления.</p></li><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.document#getfilepropertiesasync-options--callback-" target="_blank">Document.getFilePropertiesAsync</a> для доступа к свойствам файла документа, таким как URL-адрес документа.</p></li><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-" target="_blank">Document.goToByIdAsync</a> для перехода к именованным объектам и расположениям в документе.</p></li><li><p>Для надстроек области задач для Project — все методы "get" объекта <a href="https://docs.microsoft.com/javascript/api/office/office.document" target="_blank">ProjectDocument</a>. </p></li></ul>|
|**ReadAllDocument**|В дополнение к интерфейсам API, предоставляемым разрешениями **Restricted** и **ReadDocument**, дает следующие дополнительные права доступа к данным документа:<br/><ul><li><p>Методы <span class="keyword">Document.getSelectedDataAsync</span> и <span class="keyword">Document.getFileAsync</span> получают доступ к коду OOXML документа (который кроме текста может содержать форматирование, ссылки, встроенную графику, комментарии, редакции и т. д.).</p></li></ul>|
|**WriteDocument**|В дополнение к интерфейсам API, предоставляемым разрешением **Restricted**, добавляет доступ к следующим элементам API:<br/><ul><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.document#setselecteddataasync-data--options--callback-" target="_blank">Document.setSelectedDataAsync</a> для записи данных в выделенную пользователем область документа.</p></li></ul>|
|**ReadWriteDocument**|В дополнение к интерфейсам API, предоставляемым разрешениям  **Restricted**,  **ReadDocument**,  **ReadAllDocument** и **WriteDocument**, дает доступ ко всем оставшимся API, поддерживаемым контентными надстройками и надстройками области задач, в том числе методам для подписки на событий.Для доступа к этим дополнительные элементам API-интерфейса необходимо объявить разрешение  **ReadWriteDocument**:<br/><ul><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.binding#setdataasync-data--options--callback-" target="_blank">Binding.setDataAsync</a> для записи в связанных областях документа.</p></li><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding#addrowsasync-rows--options--callback-" target="_blank">TableBinding.addRowsAsync</a> для добавления строк в связанные таблицы.</p></li><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding#addcolumnsasync-tabledata--options--callback-" target="_blank">TableBinding.addColumnsAsync</a> для добавления столбцов в связанные таблицы.</p></li><li><p>Метод <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding#deletealldatavaluesasync-options--callback-" target="_blank">TableBinding.deleteAllDataValuesAsync</a> для удаления всех данных в связанной таблице.</p></li><li><p>Методы <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding#setformatsasync-cellformat--options--callback-" target="_blank">setFormatsAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding#clearformatsasync-options--callback-" target="_blank">clearFormatsAsync</a> и <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding#settableoptionsasync-tableoptions--options--callback-" target="_blank">setTableOptionsAsync</a> объекта <span class="keyword">TableBinding</span> для настройки форматирования и параметров связанных таблиц.</p></li><li><p>Все элементы объектов <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlnode" target="_blank">CustomXmlNode</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlparts" target="_blank">CustomXmlParts</a> и <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlprefixmappings" target="_blank">CustomXmlPrefixMappings</a>.</p></li><li><p>Все методы для подписки на события, поддерживаемые контентными надстройками и надстройками области задач, в частности методы <span class="keyword">addHandlerAsync</span> и <span class="keyword">removeHandlerAsync</span> объектов <a href="https://docs.microsoft.com/javascript/api/office/office.binding" target="_blank">Binding</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.document" target="_blank">Document</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.document" target="_blank">ProjectDocument</a> и <a href="https://docs.microsoft.com/javascript/api/office/office.document#settings" target="_blank">Settings</a>.</p></li></ul>|

## <a name="see-also"></a>См. также

- [Конфиденциальность и безопасность надстроек Office](../concepts/privacy-and-security.md)
    


