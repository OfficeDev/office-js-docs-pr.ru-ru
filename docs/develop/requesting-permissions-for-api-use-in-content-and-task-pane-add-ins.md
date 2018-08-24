---
title: Запрос разрешений на использование API в контентных надстройках и надстройках области задач
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: c7f303b1df20fedb41400d9b1f44512a2c5be579
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925362"
---
# <a name="requesting-permissions-for-api-use-in-content-and-task-pane-add-ins"></a>Запрос разрешений на использование API в контентных надстройках и надстройках области задач

В этой статье описываются различные уровни разрешений, которые можно объявить в манифесте контентной надстройки или надстройки области задач, чтобы указать уровень доступа JavaScript API, необходимый вашей надстройке. 




## <a name="permissions-model"></a>Модель разрешений


5-уровневая модель разрешений JavaScript API — это основа безопасности и конфиденциальности для пользователей контентных надстроек и надстроек области задач. На рис. 1 показаны пять уровней разрешений API, которые можно объявить в манифесте надстройки.


*Рис. 1. Пятиуровневая модель разрешений для контентных надстроек и надстроек области задач*

![Уровни разрешений для приложений области задач](../images/office15-app-sdk-task-pane-app-permission.png)



Эти разрешения задают подмножество API, которые контентная надстройки или надстройка области задач сможет использовать во время выполнения, когда пользователь вставляет, а затем активирует ваше приложение (доверяет ему). Чтобы объявить уровень разрешений, необходимый вашей надстройке, укажите одно из текстовых значений разрешения в элементе [Permissions](https://dev.office.com/reference/add-ins/manifest/permissions) манифеста надстройки. Следующий пример запрашивает разрешение **WriteDocument**, который разрешает использовать только методы записи в документ (но не методы чтения).




```XML
<Permissions>WriteDocument</Permissions>
```

Рекомендуется запрашивать разрешения по принципу  _минимальной привилегии_. Т. е. следует запрашивать разрешения для доступа только к минимальному подмножеству API, необходимому надстройке для правильной работы. Например, если надстройке требуется только читать данные документа пользователя, оно должно запрашивать максимум разрешение **ReadDocument**.

В следующей таблице описывается подмножество API JavaScript, предоставляемое каждым уровнем разрешений.



|**Разрешение**|**Включенное подмножество API**|
|:-----|:-----|
|**Существуют ограничения**|Методы объекта [Settings](https://dev.office.com/reference/add-ins/shared/settings) и метод [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync). Это минимальный уровень разрешений, запрашиваемый контентной надстройкой или надстройкой области задач.|
|**ReadDocument**|В дополнение к интерфейсам API, предоставляемым разрешением  **Restricted**, добавляет доступ к элементам API, необходимым для чтения документа и управления привязками.В том числе разрешается использование следующих элементов:<br/><ul><li>Метод <a href="https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync" target="_blank">Document.getSelectedDataAsync</a> для получения выбранного текста, HTML (только Word) или табличных данных, но не базовый код Open Office XML (OOXML), содержащий все данные в документе.</p></li><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/document.getfileasync" target="_blank">Document.getFileAsync</a> для получения всего текста документа, но не двоичной копии OOXML документа.</p></li><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/binding.getdataasync" target="_blank">Binding.getDataAsync</a> для чтения связанных данных в документе.</p></li><li><p>Методы <a href="https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync" target="_blank">addFromNamedItemAsync</a>, <a href="https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync" target="_blank">addFromPromptAsync</a>, <a href="https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync" target="_blank">addFromSelectionAsync</a> объекта <span class="keyword">Bindings</span> для создания привязок в документе.</p></li><li><p>Методы <a href="https://dev.office.com/reference/add-ins/shared/bindings.getallasync" target="_blank">getAllAsync</a>, <a href="https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync" target="_blank">getByIdAsync</a> и <a href="https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync" target="_blank">releaseByIdAsync</a> объекта <span class="keyword">Bindings</span> для доступа к привязкам документа и их удаления.</p></li><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync" target="_blank">Document.getFilePropertiesAsync</a> для доступа к свойствам файла документа, таким как URL-адрес документа.</p></li><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/document.gotobyidasync" target="_blank">Document.goToByIdAsync</a> для перехода к именованным объектам и расположениям в документе.</p></li><li><p>Для надстроек области задач для Project — все методы "get" объекта <a href="https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument" target="_blank">ProjectDocument</a>. </p></li></ul>|
|**ReadAllDocument**|В дополнение к интерфейсам API, предоставляемым разрешениями **Restricted** и **ReadDocument**, дает следующие дополнительные права доступа к данным документа:<br/><ul><li><p>Методы <span class="keyword">Document.getSelectedDataAsync</span> и <span class="keyword">Document.getFileAsync</span> получают доступ к коду OOXML документа (который кроме текста может содержать форматирование, ссылки, встроенную графику, комментарии, редакции и т. д.).</p></li></ul>|
|**WriteDocument**|В дополнение к интерфейсам API, предоставляемым разрешением **Restricted**, добавляет доступ к следующим элементам API:<br/><ul><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync" target="_blank">Document.setSelectedDataAsync</a> для записи данных в выделенную пользователем область документа.</p></li></ul>|
|**ReadWriteDocument**|В дополнение к интерфейсам API, предоставляемым разрешениям  **Restricted**,  **ReadDocument**,  **ReadAllDocument** и **WriteDocument**, дает доступ ко всем оставшимся API, поддерживаемым контентными надстройками и надстройками области задач, в том числе методам для подписки на событий.Для доступа к этим дополнительные элементам API-интерфейса необходимо объявить разрешение  **ReadWriteDocument**:<br/><ul><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/binding.setdataasync" target="_blank">Binding.setDataAsync</a> для записи в связанных областях документа.</p></li><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.addrowsasync" target="_blank">TableBinding.addRowsAsync</a> для добавления строк в связанные таблицы.</p></li><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.addcolumnsasync" target="_blank">TableBinding.addColumnsAsync</a> для добавления столбцов в связанные таблицы.</p></li><li><p>Метод <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.deletealldatavaluesasync" target="_blank">TableBinding.deleteAllDataValuesAsync</a> для удаления всех данных в связанной таблице.</p></li><li><p>Методы <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.setformatsasync" target="_blank">setFormatsAsync</a>, <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.clearformatsasync" target="_blank">clearFormatsAsync</a> и <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.settableoptionsasync" target="_blank">setTableOptionsAsync</a> объекта <span class="keyword">TableBinding</span> для настройки форматирования и параметров связанных таблиц.</p></li><li><p>Все элементы объектов <a href="https://dev.office.com/reference/add-ins/shared/customxmlnode.customxmlnode" target="_blank">CustomXmlNode</a>, <a href="https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts" target="_blank">CustomXmlParts</a> и <a href="https://dev.office.com/reference/add-ins/shared/customxmlprefixmappings.customxmlprefixmappings" target="_blank">CustomXmlPrefixMappings</a>.</p></li><li><p>Все методы для подписки на события, поддерживаемые контентными надстройками и надстройками области задач, в частности методы <span class="keyword">addHandlerAsync</span> и <span class="keyword">removeHandlerAsync</span> объектов <a href="https://dev.office.com/reference/add-ins/shared/binding" target="_blank">Binding</a>, <a href="https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="https://dev.office.com/reference/add-ins/shared/document" target="_blank">Document</a>, <a href="https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument" target="_blank">ProjectDocument</a> и <a href="https://dev.office.com/reference/add-ins/shared/document.settings" target="_blank">Settings</a>.</p></li></ul>|

## <a name="see-also"></a>См. также

- [Конфиденциальность и безопасность надстроек Office](../concepts/privacy-and-security.md)
    


