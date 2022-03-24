---
title: Общая объектная модель API JavaScript
description: Узнайте о Office общей объектной модели API JavaScript.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 381d3089b47fe04f403459ecae249bf68f7ca872
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743769"
---
# <a name="common-javascript-api-object-model"></a>Общая объектная модель API JavaScript

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office API JavaScript дают доступ к Office клиентского приложения. В основном такой доступ осуществляется при помощи нескольких значимых объектов. Объект [Context](#context-object) предоставляет доступ к среде выполнения после инициализации. Объект [Document](#document-object) предоставляет пользователю управление документом Excel, PowerPoint или Word. Объект [почтовых](#mailbox-object) ящиков предоставляет Outlook надстройки для сообщений, встреч и профилей пользователей. Понимание связей между этими объектами высокого уровня является основой Office надстройки.

## <a name="context-object"></a>Объект Context

**Область применения:** все типы надстроек

Когда надстройка [инициализирована](initialize-add-in.md), она содержит множество различных объектов, с которыми она может взаимодействовать в среде выполнения. Контекст среды выполнения надстройки представлен в API объектом [Context](/javascript/api/office/office.context). Объект **Context** — это основной объект, предоставляющий доступ к наиболее важным объектам API, таким как [Document](/javascript/api/office/office.document) и [Mailbox](/javascript/api/outlook/office.mailbox), которые в свою очередь предоставляют доступ к документу и содержимому почтового ящика.

Например, в надстройках области задач или контентных надстройках можно использовать свойство [document](/javascript/api/office/office.context#office-office-context-document-member) объекта **Context** для получения доступа к свойствам и методам объекта **Document**, чтобы взаимодействовать с содержимым документов Word, электронными таблицами Excel или расписаниями Project. Аналогично этому в надстройках Outlook можно использовать свойство [mailbox](/javascript/api/office/office.context#office-office-context-mailbox-member) объекта **Context** для получения доступа к свойствам и методам объекта **Mailbox**, чтобы взаимодействовать с контентом сообщений, запросов на собрание или встреч.

Объект **Context** также предоставляет доступ к свойствам [contentLanguage и displayLanguage](/javascript/api/office/office.context#office-office-context-contentlanguage-member), которые позволяет определить язык (язык), используемый в документе или элементе, или Office приложении.[](/javascript/api/office/office.context#office-office-context-displaylanguage-member) Свойство [roamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) позволяет получить доступ к элементам объекта [RoamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member), в котором хранятся настройки, специфичные для надстроек почтовых ящиков отдельных пользователей. Наконец, объект **Context** предоставляет свойство [ui](/javascript/api/office/office.context#office-office-context-ui-member), позволяющее надстройке открывать всплывающие диалоговые окна.

## <a name="document-object"></a>Объект Document

**Область применения:** надстройки области задач и контентные надстройки

Чтобы взаимодействовать с данными документа в Excel, PowerPoint и Word, интерфейс API предоставляет объект [Document](/javascript/api/office/office.document). Участники объектов могут `Document` получать доступ к данным следующим образом.

- Читать и записывать активные выделенные элементы в виде текста, смежных ячеек (матриц) или таблиц.

- Табличные данные (матрицы или таблицы).

- Привязки (созданные с помощью методов добавления объекта `Bindings` ).

- Настраиваемые части XML (только для Word).

- Параметры или состояние надстройки, сохраняемые в документе отдельными надстройками.

Объект также можно использовать для `Document` взаимодействия с данными в Project документах. Возможности API, относящиеся к Project, документированы в абстрактном классе [ProjectDocument](/javascript/api/office/office.document). Узнайте больше о создании надстроек области задач для Project в статье, посвященной [надстройкам области задач для Project](../project/project-add-ins.md).

Все эти формы доступа к данным начинаются с экземпляра абстрактного `Document` объекта.

Вы можете получить доступ к экземпляру `Document` [](/javascript/api/office/office.context#office-office-context-document-member) `Context` объекта при инициализации области задач или надстройки контента с помощью свойства документа объекта. Объект `Document` определяет общие функции доступа к данным, общие для документов Word и Excel, `CustomXmlParts` а также предоставляет доступ к объекту для документов Word.

Объект `Document` поддерживает четыре способа доступа разработчиков к содержимому документов.

- Доступ на основе выделенных фрагментов.

- Доступ на основе привязок.

- Доступ на основе настраиваемых XML-частей (только для Word).

- Доступ на основе целого документа (только для PowerPoint и Word).

Для лучшего понимания работы способов доступа к данным на основе выделенных фрагментов и привязок мы сначала объясним, как API доступа к данным обеспечивают единообразный доступ к данным в различных приложениях Office.

### <a name="consistent-data-access-across-office-applications"></a>Единообразный доступ к данным в приложениях Office

 **Область применения:** надстройки области задач и контентные надстройки

Чтобы создать расширения, которые легко работают в разных Office документах, API javaScript Office отбирает особенности каждого приложения Office с помощью общих типов данных и возможность принудительного принуждения содержимого различных документов к трем общим типам данных.

#### <a name="common-data-types"></a>Общие типы данных

Во время доступа к данным как через выделенные фрагменты, так и через привязки контент документа предоставляется через типы данных, которые являются общими во всех поддерживаемых приложениях Office. В Office 2013 г. поддерживаются три основных типа данных.

|**Тип данных**|**Описание**|**Поддержка ведущего приложения**|
|:-----|:-----|:-----|
|Текст|Предоставляет строковое представление данных в выделенном фрагменте или привязке.|В Excel 2013, Project 2013 и PowerPoint 2013 поддерживается только обычный текст. В Word 2013 поддерживаются три текстовых формата: обычный текст, HTML и Office Open XML (OOXML). При выборе текста в ячейке в Excel методы выделения осуществляют чтение и запись всего содержимого ячейки, даже если в ячейке выделена только часть текста. При выделении текста в Word и PowerPoint методы выделения осуществляют чтение и запись только для выполнения выбранных символов. Project 2013 и PowerPoint 2013 поддерживает только доступ к данным на основе выделения.|
|Матрица|Предоставляет данные в выборе или привязке как двумерный объект **Array**, который в JavaScript реализован как массив массивов. Например, две строки значений **string** в двух столбцах будут выглядеть как ` [['a', 'b'], ['c', 'd']]`, а один столбец, состоящий из трех строк, — как `[['a'], ['b'], ['c']]`.|Доступ к матричным данным поддерживается только в Excel 2013 и Word 2013.|
|Таблица|Предоставляет данные в выделенном фрагменте или привязке в виде объекта [TableData](/javascript/api/office/office.tabledata). Объект `TableData` предоставляет данные через свойства `headers` и `rows` свойства.|Доступ к данным таблицы поддерживается только в Excel 2013 и Word 2013.|

#### <a name="data-type-coercion"></a>Приведение типов данных

Методы доступа к `Document` данным для объектов и [](/javascript/api/office/office.binding) привязки поддерживают указание нужного типа данных с помощью параметра _coercionType_ этих методов и соответствующих значений [принужденияType](/javascript/api/office/office.coerciontype). Вне зависимости от действительной формы привязки различные приложения Office поддерживают общие типы данных, пытаясь привести данные в запрошенный тип данных. Например, если выделена таблица Word или абзац, разработчик может считывать эту таблицу в виде неформатированного текста, HTML, Office Open XML или таблицы, а API производит необходимые преобразования данных.

> [!TIP]
> **В каких случаях следует использовать для доступа к данным матрицы, а в каких — coercionType?** Если при добавлении строк и столбцов необходимо динамическое развитие табулярных данных, и вам необходимо работать с загонами таблицы, следует использовать тип данных таблицы (указав параметр _coercionType_ `Document` `Binding` `"table"` `Office.CoercionType.Table`метода доступа к объектным данным как или ). Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными. Если вы не планируете добавлять строки и столбцы, а ваши данные не требуют функции загона, следует использовать тип матричных данных (указав параметр  _coercionType_ `"matrix"` `Office.CoercionType.Matrix`метода доступа к данным как или), который обеспечивает более простую модель взаимодействия с данными.

Если данные невозможно привести к заданному типу, то свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) в функции обратного вызова возвращает значение `"failed"`, и можно использовать свойство [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member), чтобы получить доступ к объекту [Error](/javascript/api/office/office.error) со сведениями о причине ошибки во время вызова метода.

## <a name="work-with-selections-using-the-document-object"></a>Работа с выборами с помощью объекта Document

Объект `Document` предоставляет методы, которые позволяют читать и записывать текущий выбор пользователя в "получить и установить" моды. Для этого объект предоставляет `Document` средства и `getSelectedDataAsync` методы `setSelectedDataAsync` .

Примеры кода, демонстрирующие выполнение задач с выделенными фрагментами, см. в статье [Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).

## <a name="work-with-bindings-using-the-bindings-and-binding-objects"></a>Работа с привязками с помощью объектов Bindings и Binding

Доступ к данным на основе привязок позволяет надстройкам области задач и контентным надстройкам получать единообразный доступ к определенной области документа или электронной таблицы через идентификатор, связанный с привязкой. Сначала надстройка должна создать привязку с помощью вызова одного из методов, связывающих часть документа с уникальным идентификатором: [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)), [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) или [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)). После настройки привязки надстройка может использовать предоставленный идентификатор для доступа к данным, содержащимся в связанном регионе документа или электронной таблицы. Создание привязки обеспечивает следующее значение для надстройки.

- Разрешает доступ к общим структурам данных в поддерживаемых приложениях Office, таким как: таблицы, диапазоны или текст (связанная последовательность знаков).

- Позволяет производить операции чтения или записи без необходимости выделения пользователем фрагмента.

- Устанавливает отношение между надстройкой и данными в документе. Привязки сохраняются в документе и могут использоваться позже.

Установка привязки также позволяет подписываться на данные и выбирать изменения событий, относящиеся к конкретной области документа или электронной таблицы. Это означает, что надстройка уведомляется только об изменениях, происходящих внутри данной конкретной области, в отличие от изменений, затрагивающих в целом весь документ или электронную таблицу.

Объект [Bindings](/javascript/api/office/office.bindings) предоставляет метод [getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)), который обеспечивает доступ к набору всех привязок, установленных в этом документе или листе. Доступ к отдельной привязке можно получить по ее идентификатору с помощью методов [Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) или [Office.select](/javascript/api/office). Вы можете `Bindings` установить новые привязки, а также удалить существующие, используя один из следующих методов объекта: [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)), [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)), [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)) или [releaseByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-releasebyidasync-member(1)).

Существует три различных типа привязки, которые вы указываете с параметром  _bindingType_ `addFromSelectionAsync`при создании привязки с методами или методами `addFromPromptAsync` `addFromNamedItemAsync` .

|**Тип привязки**|**Описание**|**Поддержка ведущего приложения**|
|:-----|:-----|:-----|
|Привязка текста|Выполняет привязку к области документа, которая может быть представлена как текст.|В Word поддерживается большинство связанных выделений, тогда как в Excel для привязки текста можно использовать только выделения отдельных ячеек. Excel поддерживает только обычный текст, а Word — три формата: обычный текст, HTML и Open XML для Office.|
|Привязка матрицы|Выполняет привязку к фиксированной области документа, содержащей табличные данные без заголовков. Данные в привязке матрицы записываются или считываются как двумерный **Array**, который в JavaScript реализован как массив массивов. Например, две строки значений **string** в двух столбцах можно записать или прочитать как ` [['a', 'b'], ['c', 'd']]`, а один столбец, состоящий из трех строк, — как `[['a'], ['b'], ['c']]`.|В Excel для установки матричной привязки может использоваться любое связанное выделение ячеек. В Word матричная привязка поддерживается только таблицами.|
|Привязка таблицы|Выполняет привязку к области документа, содержащей таблицу с заголовками. Данные в привязке таблицы записываются или считываются как объект [TableData](/javascript/api/office/office.tabledata). Объект `TableData` предоставляет данные с помощью свойств **заглавных и** **строк** .|Любая таблица Excel или Word может быть основой для табличной привязки. После создания табличной привязки каждая новая строка или столбец, добавляемые пользователем в таблицу, автоматически включаются в привязку. |

<br/>

После создания привязки с помощью одного из трех методов добавления `Bindings` объекта можно работать с данными и свойствами привязки с помощью методов соответствующего объекта: [MatrixBinding](/javascript/api/office/office.matrixbinding), [TableBinding](/javascript/api/office/office.tablebinding) или [TextBinding](/javascript/api/office/office.textbinding). Все три этих объекта наследуют [методы getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) и [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)) `Binding` объекта, которые позволяют взаимодействовать с связанными данными.

Примеры кода, демонстрирующие выполнение задач с привязками, см. в статье [Привязка к областям в документе или электронной таблице](bind-to-regions-in-a-document-or-spreadsheet.md).

## <a name="work-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Работа с пользовательскими частями XML с помощью объектов CustomXmlParts и CustomXmlPart

 **Область применения:** надстройки области задач Word

Объекты [CustomXmlParts](/javascript/api/office/office.customxmlparts) и [CustomXmlPart](/javascript/api/office/office.customxmlpart) интерфейса API предоставляют доступ к настраиваемым частям XML в документах Word, которые позволяют работать с содержимым документа на основе XML. Демонстрации работы с `CustomXmlParts` `CustomXmlPart` объектами и объектами см. в примере [кода Word-add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) .

## <a name="work-with-the-entire-document-using-the-getfileasync-method"></a>Работа со всем документом с помощью метода getFileAsync

 **Область применения:** надстройки области задач Word и PowerPoint

Метод [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) и члены объектов [File](/javascript/api/office/office.file) и [Slice](/javascript/api/office/office.slice) предоставляют возможность получения целого файла документа Word и PowerPoint в виде порций (блоков) размером до 4 МБ. Дополнительные сведения см. в статье [Получение всего документа из надстройки PowerPoint или Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="mailbox-object"></a>Объект Mailbox

**Область применения:** надстройки Outlook

Надстройки Outlook, в основном, используют набор API, предоставляемый через объект [Mailbox](/javascript/api/outlook/office.mailbox). Чтобы получить объекты и члены специально для использования в надстройках Outlook, такие как объект [Item](/javascript/api/outlook/office.item), используйте свойство [mailbox](/javascript/api/office/office.context#office-office-context-mailbox-member) объекта **Context** для получения доступа к объекту **Mailbox**, как показано в следующей строке кода.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Кроме того, Outlook надстройки могут использовать следующие объекты.

- `Office` объект: для инициализации.

- `Context` объект: для доступа к содержимому и отображения языковых свойств.

- `RoamingSettings`объект: для Outlook настраиваемые параметры надстройки в почтовом ящике пользователя, где установлена надстройка.

Сведения об использовании JavaScript в надстройках Outlook см. в статье [Надстройки Outlook](../outlook/outlook-add-ins-overview.md).
