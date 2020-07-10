---
title: Автоматическое открытие области задач с документом
description: Узнайте, как настроить надстройку Office для автоматического открытия при открытии документа.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 85b421a569ccb83c3d07f0f10fd4767929332f96
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093709"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>Автоматическое открытие области задач с документом

Вы можете использовать команды надстроек в надстройке Office для расширения пользовательского интерфейса Office, добавляя кнопки на ленту приложения Office. Когда пользователи нажимают кнопки, выполняются различные операции (например, открывается область задач).

В некоторых сценариях требуется, чтобы область задач открывалась автоматически вместе с документом без явного взаимодействия с пользователем. Функция автоматического открытия области задач, представленная в наборе требований AddInCommands 1.1, позволяет автоматически открывать область задач, если это требуется в вашем сценарии.


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>Чем функция автоматического открытия отличается от вставки области задач?

Если пользователь запускает надстройки, которые не используют команды надстроек (например, в Office 2013), они вставляются в документ и сохраняются в нем. Таким образом, при открытии этого документа другими пользователями, им будет предложено установить надстройку и откроется область задач. Проблема с этой моделью состоит в том, что во многих случаях пользователи не хотят, чтобы надстройка сохранялась в документе. Например, учащийся, который использует надстройку словаря в документе Word, может не захотеть, чтобы его преподаватели или одноклассники получили запрос на установку этой надстройки при открытии документа.

Функция автоматического открытия позволяет явно определить или дать разрешение пользователю определять необходимость в сохранении конкретной надстройки области задач в конкретном документе.

## <a name="support-and-availability"></a>Поддержка и доступность

Функция автоматического открытия <!-- in **developer preview** and it is only --> поддерживается в продуктах и платформах, перечисленных ниже.

|**Продукты**|**Платформы**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|Поддерживаемые платформы для всех продуктов:<ul><li>Office on Windows Desktop. Build 16.0.8121.1000+</li><li>Office on Mac. Build 15.34.17051500+</li><li>Office в Интернете</li></ul>|


## <a name="best-practices"></a>Рекомендации

При использовании функции автоматического открытия придерживайтесь указанных рекомендаций.

- Используйте функцию автоматического открытия, если она повысит эффективность работы пользователей в подобных случаях:
  - When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.
  - When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.
- Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.  
- Обнаружение набора требований используется для определения доступности функции автоматического открытия и обеспечения резервного поведения, если это не так.
- Не используйте функцию автоматического открытия, чтобы искусственно увеличивать показатели использования надстройки. Если вы не хотите, чтобы ваша надстройка автоматически открывалась с определенными документами, эта функция может раздражать пользователей.

    > [!NOTE]
    > Если корпорация Майкрософт обнаружит, что функция автоматического открытия применяется не по назначению, возможно исключение вашей надстройки из AppSource.

- Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.  

## <a name="implementation"></a>Реализация

Выполните следующие действия, чтобы использовать функцию автоматического открытия.

- Укажите область задач, которую необходимо открывать автоматически.
- Отметьте документ, в котором будет автоматически открываться эта область задач.

> [!IMPORTANT]
> The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.

### <a name="step-1-specify-the-task-pane-to-open"></a>Этап 1. Указание области задач, которую необходимо открывать

To specify the task pane to open automatically, set the [TaskpaneId](../reference/manifest/action.md#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.

Ниже представлен пример, где для TaskPaneId задано значение Office.AutoShowTaskpaneWithDocument.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>Этап 2. Установка отметки для документа, вместе с которым будет автоматически открываться область задач

You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.  


#### <a name="tag-the-document-on-the-client-side"></a>Установка отметки для документа на стороне клиента

Используйте метод Office.js [settings.set](/javascript/api/office/office.settings), чтобы установить для **Office.AutoShowTaskpaneWithDocument** значение **true**, как показано в следующем примере.

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

Используйте этот метод, если нужно отметить документ в рамках взаимодействия с надстройкой (например, после создания пользователем привязки или выбора параметра, он сможет указать необходимость в автоматическом открытии области).

#### <a name="use-open-xml-to-tag-the-document"></a>Установка отметки для документа с помощью Open XML

You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).

Добавьте в документ две части Open XML:

- часть `webextension`.
- часть `taskpane`.

В примере ниже показано, как добавить часть `webextension`.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
   <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

Часть `webextension` содержит контейнер свойств, а также свойство под названием **Office.AutoShowTaskpaneWithDocument**, для которого необходимо установить значение `true`.

Часть `webextension` также содержит ссылку на магазин или каталог с атрибутами для `id`, `storeType`, `store` и `version`. Только четыре значения `storeType` относятся к функции автоматического открытия. Значения остальных трех атрибутов зависят от значения для `storeType`, как показано в таблице ниже.

| Значение **`storeType`** | Значение **`id`**    |**Значение `store`** | **Значение `version`**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|ИД ресурса AppSource для надстройки (см. примечание)|Код языка для AppSource (например, "ru-ru").|Версия каталога AppSource (см. примечание)|
|FileSystem (сетевая папка)|GUID надстройки в ее манифесте.|Путь к сетевой папке (например, "\\\\MyComputer\\MySharedFolder").|Версия в манифесте надстройки.|
|EXCatalog (развертывание через Exchange Server) |GUID надстройки в ее манифесте.|EXCatalog. Строка EXCatalog — это строка, используемая для надстроек, использующих централизованное развертывание в центре администрирования Microsoft 365.|Версия в манифесте надстройки.
|Registry (реестр системы)|GUID надстройки в ее манифесте.|"developer"|Версия в манифесте надстройки.|

> [!NOTE]
> To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.

Дополнительные сведения об исправлении webextension см. в документе [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).

В примере ниже показано, как добавить часть `taskpane`.

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

Обратите внимание, что в этом примере для атрибута `visibility` установлено значение "0". Это означает, что после добавления частей webextension и `taskpane` при первом открытии документа пользователю необходимо будет установить надстройку, нажав кнопку **Надстройка** на ленте. После этого область задач надстройки будет открываться автоматически вместе с файлом. Кроме того, если установить для `visibility` значение "0", можно с помощью Office.js предоставить пользователям возможность включать и выключать функцию автоматического открытия. В частности, ваш скрипт устанавливает для параметра документа **Office.AutoShowTaskpaneWithDocument** значение `true` или `false`. (Дополнительные сведения см. в разделе [Установка отметки для документа на стороне клиента](#tag-the-document-on-the-client-side).)

If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.

Значение "1" отлично подходит для свойства `visibility`, если надстройка и шаблон или содержимое документа интегрированы настолько тесно, что пользователь не откажется от использования функции автоматического открытия.

> [!NOTE]
> If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.

An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.

## <a name="test-and-verify-opening-task-panes"></a>Тестирование и проверка открытия областей задач

Вы можете развернуть тестовую версию надстройки, которая будет автоматически открывать область задач с помощью централизованного развертывания с помощью центра администрирования Microsoft 365. В примере ниже показано, как надстройки вставляются в каталог централизованного развертывания при помощи EXCatalog (версии из магазина).

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

Вы можете протестировать предыдущий пример, используя подписку на Microsoft 365, чтобы испытать централизованное развертывание и убедиться в том, что ваша надстройка работает должным образом. Если у вас еще нет подписки на Microsoft 365, вы можете получить бесплатную, 90 день реневабле подписку на Microsoft 365, присоединяясь к [программе microsoft 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## <a name="see-also"></a>См. также

Пример использования функции автоматического открытия см. [на странице с примерами команд для надстройки Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).
[Присоединяйтесь к программе для разработчиков Microsoft 365](/office/developer-program/office-365-developer-program).
