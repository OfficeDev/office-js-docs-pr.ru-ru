---
title: Запрос разрешений на использование API в надстройках
description: Сведения о различных уровнях разрешений для объявления в манифесте контентной надстройки или надстройки области задач, чтобы указать уровень доступа к API JavaScript.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: f2a4fbcc6e1f3aa90b0a54e5fc3178e73c00e0ae
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810318"
---
# <a name="requesting-permissions-for-api-use-in-add-ins"></a>Запрос разрешений на использование API в надстройках

В этой статье описываются различные уровни разрешений, которые можно объявить в манифесте контентной надстройки или надстройки области задач, чтобы указать уровень доступа JavaScript API, необходимый вашей надстройке.

> [!NOTE]
> Сведения об уровнях разрешений для почтовых надстроек (Outlook) см. в статье [Модель разрешений Outlook](../outlook/privacy-and-security.md#permissions-model).

## <a name="permissions-model"></a>Модель разрешений

5-уровневая модель разрешений JavaScript API — это основа безопасности и конфиденциальности для пользователей контентных надстроек и надстроек области задач. На рис. 1 показаны пять уровней разрешений API, которые можно объявить в манифесте надстройки.

*Рис. 1. Пятиуровневая модель разрешений для контентных надстроек и надстроек области задач*

![Уровни разрешений для приложений области задач.](../images/office15-app-sdk-task-pane-app-permission.png)

Эти разрешения указывают подмножество API, которое [среда выполнения](../testing/runtimes.md) надстройки позволит использовать контентной надстройке или надстройке области задач при вставке пользователем, а затем активирует (доверяет) надстройку. Чтобы объявить уровень разрешений, необходимый вашей надстройке, укажите одно из текстовых значений разрешения в элементе [Permissions](/javascript/api/manifest/permissions) манифеста надстройки. Следующий пример запрашивает разрешение **WriteDocument**, который разрешает использовать только методы записи в документ (но не методы чтения).

```XML
<Permissions>WriteDocument</Permissions>
```

As a best practice, you should request permissions based on the principle of  _least privilege_. That is, you should request permission to access only the minimum subset of the API that your add-in requires to function correctly. For example, if your add-in needs only to read data in a user's document for its features, you should request no more than the **ReadDocument** permission.

В следующей таблице описывается подмножество API JavaScript, предоставляемое каждым уровнем разрешений.

|**Разрешение**|**Включенное подмножество API**|
|:-----|:-----|
|**Restricted**|Методы объекта [Settings](/javascript/api/office/office.settings) и метод [Document.getActiveViewAsync](/javascript/api/office/office.document#office-office-document-getactiveviewasync-member(1)).Это минимальный уровень разрешений, запрашиваемый контентной надстройкой или надстройкой области задач.|
|**ReadDocument**|Помимо API, разрешенного **ограниченным** разрешением, добавляет доступ к членам API, необходимым для чтения документа и управления привязками. Сюда входит использование следующих средств:<br/><ul><li>Метод <a href="/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_" target="_blank">Document.getSelectedDataAsync</a> для получения выбранного текста, HTML (только Word) или табличных данных, но не базовый код Open Office XML (OOXML), содержащий все данные в документе.</p></li><li><p>Метод <a href="/javascript/api/office/office.document#getFileAsync_fileType__options__callback_" target="_blank">Document.getFileAsync</a> для получения всего текста документа, но не двоичной копии OOXML документа.</p></li><li><p>Метод <a href="/javascript/api/office/office.binding#getDataAsync_options__callback_" target="_blank">Binding.getDataAsync</a> для чтения связанных данных в документе.</p></li><li><p>Методы <a href="/javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_" target="_blank">addFromNamedItemAsync</a>, <a href="/javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_" target="_blank">addFromPromptAsync</a>, <a href="/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_" target="_blank">addFromSelectionAsync</a> объекта <span class="keyword">Bindings</span> для создания привязок в документе.</p></li><li><p>Методы <a href="/javascript/api/office/office.bindings#getAllAsync_options__callback_" target="_blank">getAllAsync</a>, <a href="/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_" target="_blank">getByIdAsync</a> и <a href="/javascript/api/office/office.bindings#releaseByIdAsync_id__options__callback_" target="_blank">releaseByIdAsync</a> объекта <span class="keyword">Bindings</span> для доступа к привязкам документа и их удаления.</p></li><li><p>Метод <a href="/javascript/api/office/office.document#getFilePropertiesAsync_options__callback_" target="_blank">Document.getFilePropertiesAsync</a> для доступа к свойствам файла документа, таким как URL-адрес документа.</p></li><li><p>Метод <a href="/javascript/api/office/office.document#goToByIdAsync_id__goToType__options__callback_" target="_blank">Document.goToByIdAsync</a> для перехода к именованным объектам и расположениям в документе.</p></li><li><p>Для надстроек области задач для Project — все методы "get" объекта <a href="/javascript/api/office/office.document" target="_blank">ProjectDocument</a>. </p></li></ul>|
|**ReadAllDocument**|В дополнение к API, разрешенным разрешениями **Restricted** и **ReadDocument** , предоставляет следующий дополнительный доступ к данным документа.<br/><ul><li><p>Методы <span class="keyword">Document.getSelectedDataAsync</span> и <span class="keyword">Document.getFileAsync</span> получают доступ к коду OOXML документа (который кроме текста может содержать форматирование, ссылки, встроенную графику, комментарии, редакции и т. д.).</p></li></ul>|
|**WriteDocument**|Помимо API, разрешенного **ограниченным** разрешением, добавляет доступ к следующим членам API.<br/><ul><li><p>Метод <a href="/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_" target="_blank">Document.setSelectedDataAsync</a> для записи данных в выделенную пользователем область документа.</p></li></ul>|
|**ReadWriteDocument**|Помимо API, разрешенного разрешениями **Restricted**, **ReadDocument**, **ReadAllDocument** и **WriteDocument** , включает доступ ко всем остальным API, поддерживаемым контентными надстройками и надстройками области задач, включая методы подписки на события. Чтобы получить доступ к дополнительным членам API, необходимо объявить разрешение **ReadWriteDocument** :<br/><ul><li><p>Метод <a href="/javascript/api/office/office.binding#setDataAsync_data__options__callback_" target="_blank">Binding.setDataAsync</a> для записи связанных областей документа.</p></li><li><p>Метод <a href="/javascript/api/office/office.tablebinding#addRowsAsync_rows__options__callback_" target="_blank">TableBinding.addRowsAsync</a> для добавления строк в связанные таблицы.</p></li><li><p>Метод <a href="/javascript/api/office/office.tablebinding#addColumnsAsync_tableData__options__callback_" target="_blank">TableBinding.addColumnsAsync</a> для добавления столбцов в связанные таблицы.</p></li><li><p>Метод <a href="/javascript/api/office/office.tablebinding#deleteAllDataValuesAsync_options__callback_" target="_blank">TableBinding.deleteAllDataValuesAsync</a> для удаления всех данных в связанной таблице.</p></li><li><p>Методы <a href="/javascript/api/office/office.tablebinding#setFormatsAsync_cellFormat__options__callback_" target="_blank">setFormatsAsync</a>, <a href="/javascript/api/office/office.tablebinding#clearFormatsAsync_options__callback_" target="_blank">clearFormatsAsync</a> и <a href="/javascript/api/office/office.tablebinding#setTableOptionsAsync_tableOptions__options__callback_" target="_blank">setTableOptionsAsync</a> объекта <span class="keyword">TableBinding</span> для настройки форматирования и параметров связанных таблиц.</p></li><li><p>Все элементы объектов <a href="/javascript/api/office/office.customxmlnode" target="_blank">CustomXmlNode</a>, <a href="/javascript/api/office/office.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="/javascript/api/office/office.customxmlparts" target="_blank">CustomXmlParts</a> и <a href="/javascript/api/office/office.customxmlprefixmappings" target="_blank">CustomXmlPrefixMappings</a>.</p></li><li><p>Все методы для подписки на события, поддерживаемые контентными надстройками и надстройками области задач, в частности методы <span class="keyword">addHandlerAsync</span> и <span class="keyword">removeHandlerAsync</span> объектов <a href="/javascript/api/office/office.binding" target="_blank">Binding</a>, <a href="/javascript/api/office/office.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="/javascript/api/office/office.document" target="_blank">Document</a>, <a href="/javascript/api/office/office.document" target="_blank">ProjectDocument</a> и <a href="/javascript/api/office/office.document#settings" target="_blank">Settings</a>.</p></li></ul>|

## <a name="see-also"></a>См. также

- [Конфиденциальность и безопасность надстроек Office](../concepts/privacy-and-security.md)
