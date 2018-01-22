
# <a name="understanding-the-javascript-api-for-office"></a>Общие сведения об интерфейсе JavaScript API для Office



В этой статье можно узнать об интерфейсе API JavaScript для Office и о том, как его использовать. Справочные сведения см. в разделе [API JavaScript для Office](http://dev.office.com/reference/add-ins/javascript-api-for-office). О том, как обновить файлы проекта Visual Studio до последней версии API JavaScript для Office, см. в статье [Обновление версии API JavaScript для Office и файлов схемы манифеста](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

>
  **Примечание.** Если вы планируете [публиковать](../publish/publish.md) надстройку в Магазине Office, она должна соответствовать [политикам проверки Магазина Office](https://msdn.microsoft.com/en-us/library/jj220035.aspx), чтобы пройти проверку. Например, работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) и на [странице с информацией о доступности и ведущих приложениях для надстроек Office](https://dev.office.com/add-in-availability).

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Ссылки на библиотеку API JavaScript для Office в надстройке

Библиотека [API JavaScript для Office](http://dev.office.com/reference/add-ins/javascript-api-for-office) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например Excel-15.js и Outlook-15.js. Простейший способ сослаться на API — использовать нашу сеть доставки содержимого (CDN), добавив следующий код `<script>` в тег `<head>` страницы:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Это приведет к скачиванию и кэшированию файлов JavaScript API для Office при первой загрузке надстройки, чтобы убедиться, что она использует актуальную реализацию Office.js и сопутствующих файлов для указанной версии.

Подробные сведения о CDN Office.js, включая способы управления версиями и обратной совместимостью, см. в статье [Указание ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Инициализация надстройки


 **Область применения:** все типы надстроек


Библиотека Office.js включает событие инициализации, которое вызывается, когда API полностью загружено и готово к взаимодействию с пользователем. С помощью обработчика события **initialize** можно реализовать распространенные сценарии инициализации надстройки, например предложение выбрать некоторые ячейки в Excel и вставку диаграммы, инициализированной с помощью выбранных значений. Кроме того, с помощью обработчика события initialize можно инициализировать другую пользовательскую логику для надстройки, например установку привязок, запросы на значения параметров надстройки по умолчанию и т. д.

 В простейшем случае событие initialize будет выглядеть как в следующем примере:     

```js
Office.initialize = function () { };
```
При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, их следует размещать внутри события Office.initialize. Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` должна выглядеть следующим образом:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```
На всех страницах надстроек Office необходимо назначить обработчик события initialize, **Office.initialize**. Если не назначить обработчик события, при запуске надстройки может возникнуть ошибка. Кроме того, если пользователь попробует использовать надстройку с веб-клиентом Office Online, например Excel Online, PowerPoint Online или Outlook Web App, произойдет сбой. Если вам не нужен код инициализации, то функция, назначенная событию **Office.initialize**, может не содержать кода, как показано в первом из приведенных выше примеров.

Дополнительные сведения о последовательности событий при инициализации надстройки см. в статье [Загрузка модели DOM и среды выполнения](../../docs/develop/loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Причина инициализации
Для надстроек области задач и контентных надстроек Office.initialize обеспечивает дополнительный параметр _reason_. Этот параметр можно использовать для определения способа, каким надстройка была добавлена в текущий документ. Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые или когда она уже существует в документе. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
      switch (reason) {
        case 'inserted': console.log('The add-in was just inserted.');
        case 'documentOpened': console.log('The add-in is already part of the document.');
    }
}
```
Дополнительные сведения см. в статьях [Событие Office.initialize Event](http://dev.office.com/reference/add-ins/shared/office.initialize) и [Перечисление InitializationReason](http://dev.office.com/reference/add-ins/shared/initializationreason-enumeration). 

## <a name="context-object"></a>Объект Context

 **Область применения:** все типы надстроек

Когда надстройка инициализирована, она содержит множество различных объектов, с которыми она может взаимодействовать в среде выполнения. Контекст среды выполнения надстройки представлен в API объектом [Context](http://dev.office.com/reference/add-ins/shared/office.context). Объект **Context** — это основной объект, предоставляющий доступ к наиболее важным объектам API, таким как [Document](http://dev.office.com/reference/add-ins/shared/document) и [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox), которые в свою очередь предоставляют доступ к документу и содержимому почтового ящика.

Например, в надстройках области задач или контентных надстройках можно использовать свойство [document](http://dev.office.com/reference/add-ins/shared/office.context.document) объекта **Context** для получения доступа к свойствам и методам объекта **Document**, чтобы взаимодействовать с содержимым документов Word, электронными таблицами Excel или расписаниями Project. Аналогично этому в надстройках Outlook можно использовать свойство [mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) объекта **Context** для получения доступа к свойствам и методам объекта **Mailbox**, чтобы взаимодействовать с контентом сообщений, запросов на собрание или встреч.

Объект **Context** также предоставляет доступ к свойствам [contentLanguage](http://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) и [displayLanguage](http://dev.office.com/reference/add-ins/shared/office.context.displaylanguage), которые позволяют задать языковые параметры, используемые в документе, элементе или в основном приложении, и к свойству [roamingSettings](http://dev.office.com/reference/add-ins/outlook/Office.context), позволяющему получать доступ к элементам объекта [RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings). Наконец, объект **Context** предоставляет свойство [ui](http://dev.office.com/reference/add-ins/shared/officeui), позволяющее надстройке открывать всплывающие диалоговые окна.


## <a name="document-object"></a>Объект Document


 **Область применения:** надстройки области задач и контентные надстройки

Чтобы взаимодействовать с данными документа в Excel, PowerPoint и Word, интерфейс API предоставляет объект [Document](http://dev.office.com/reference/add-ins/shared/document). Члены объекта **Document** можно использовать для получения доступа к данным в следующем виде:


- Читать и записывать активные выделенные элементы в виде текста, смежных ячеек (матриц) или таблиц.
    
- Табличные данные (матрицы или таблицы).
    
- Привязки (созданные методами add объекта **Bindings**).
    
- Настраиваемые части XML (только для Word).
    
- Параметры или состояние надстройки, сохраняемые в документе отдельными надстройками.
    
Кроме того, объект **Document** можно использовать для взаимодействия с данными в документах Project. Возможности API, относящиеся к Project, документированы в абстрактном классе [ProjectDocument](http://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument). Узнайте больше о создании надстроек области задач для Project в статье, посвященной [надстройкам области задач для Project](../project/project-add-ins.md).

Все эти виды доступа к данным начинаются с экземпляра абстрактного объекта **Document**.

Получить доступ к экземпляру объекта **Document** после инициализации надстройки области задач или контентной надстройки можно с помощью свойства [document](http://dev.office.com/reference/add-ins/shared/office.context.document) объекта **Context**. Объект **Document** определяет общие функции доступа к данным, используемые в документах Word и Excel, а также предоставляет доступ к объекту **CustomXmlParts** для документов Word.

Объект **Document** поддерживает четыре способа, которые разработчики могут применить для доступа к контенту документов:


- Доступ на основе выделенных фрагментов.
    
- Доступ на основе привязок.
    
- Доступ на основе настраиваемых XML-частей (только для Word).
    
- Доступ на основе целого документа (только для PowerPoint и Word).
    
Для лучшего понимания работы способов доступа к данным на основе выделенных фрагментов и привязок мы сначала объясним, как API доступа к данным обеспечивают единообразный доступ к данным в различных приложениях Office.


### <a name="consistent-data-access-across-office-applications"></a>Единообразный доступ к данным в приложениях Office

 **Область применения:** надстройки области задач и контентные надстройки

Чтобы создать расширения, которые прозрачно работают в различных документах Office, абстрактные классы API JavaScript для Office исключают особенности конкретных приложений Office с помощью общих типов данных и возможности приводить содержимое разных документов к трем общим типам данных.


#### <a name="common-data-types"></a>Общие типы данных

Во время доступа к данным как через выделенные фрагменты, так и через привязки контент документа предоставляется через типы данных, которые являются общими во всех поддерживаемых приложениях Office. В Office 2013 поддерживается три основных типа данных:



|**Тип данных**|**Описание**|**Поддержка ведущего приложения**|
|:-----|:-----|:-----|
|Текст|Предоставляет строковое представление данных в выделенном фрагменте или привязке|В Excel 2013, Project 2013 и PowerPoint 2013 поддерживается только обычный текст. В Word 2013 поддерживаются три текстовых формата: обычный текст, HTML и Office Open XML (OOXML). При выборе текста в ячейке в Excel методы выделения осуществляют чтение и запись всего содержимого ячейки, даже если в ячейке выделена только часть текста. При выделении текста в Word и PowerPoint методы выделения осуществляют чтение и запись только для выполнения выбранных символов. Project 2013 и PowerPoint 2013 поддерживает только доступ к данным на основе выделения.|
|Матрица|Предоставляет данные в выборе или привязке как двумерный массив **Array**, который в JavaScript реализован как массив массивов. Например, две строки значений **string** в двух столбцах будут выглядеть как ` [['a', 'b'], ['c', 'd']]`, а один столбец, состоящий из трех строк, — как `[['a'], ['b'], ['c']]`.|Доступ к данным на основе матрицы поддерживается только в Excel 2013 и Word 2013.|
|Таблица|Предоставляет данные в выделенном фрагменте или привязке в виде объекта [TableData](http://dev.office.com/reference/add-ins/shared/tabledata). Объект **TableData** предоставляет данные через свойства **headers** и **rows**.|Доступ к данным таблицы поддерживается только в Excel 2013 и Word 2013.|

#### <a name="data-type-coercion"></a>Приведение типов данных

Методы доступа к данным объектам **Document** и [Binding](http://dev.office.com/reference/add-ins/shared/binding) поддерживают указание желаемого типа данных, используя параметр _coercionType_ этих методов и соответствующие значения перечисления [CoercionType](http://dev.office.com/reference/add-ins/shared/coerciontype-enumeration). Вне зависимости от действительной формы привязки различные приложения Office поддерживают общие типы данных, пытаясь привести данные в запрошенный тип данных. Например, если выделена таблица Word или абзац, разработчик может указать чтение этой таблицы в виде неформатированного текста, HTML, Office Open XML или таблицы, а API производит необходимые преобразования данных.


 >**Совет.** **В каких случаях следует использовать для доступа к данным матрицы, а в каких — coercionType?** Если вы хотите, чтобы табличные данные динамически росли при добавлении строк и столбцов, и вам нужно работать с заголовками таблиц, используйте табличный тип данных (для этого укажите параметр _coercionType_ метода для доступа к данным объекта **Document** либо **Binding** в виде `"table"` или **Office.CoercionType.Table**). Добавление строк и столбцов в структуру данных поддерживается как для табличных, так и для матричных данных, но добавление строк и столбцов в конец поддерживается только для табличных данных. Если вы не планируете добавлять строки и столбцы, а для данных не требуются заголовки, следует использовать матричный тип данных (указав параметр _coercionType_ метода доступа к данным в виде `"matrix"` или **Office.CoercionType.Matrix**), что позволяет использовать упрощенный способ взаимодействия с данными.

Если данные невозможно привести к заданному типу, то свойство [AsyncResult.status](http://dev.office.com/reference/add-ins/shared/asyncresult.error) в функции обратного вызова возвращает значение `"failed"`, и можно использовать свойство [AsyncResult.error](http://dev.office.com/reference/add-ins/shared/asyncresult.context), чтобы получить доступ к объекту [Error](http://dev.office.com/reference/add-ins/shared/error) со сведениями о причине ошибки во время вызова метода.


## <a name="working-with-selections-using-the-document-object"></a>Работа с выделенными фрагментами с помощью объекта Document


Объект **Document** предоставляет методы, позволяющие выполнять чтение и запись к текущему выделенному фрагменту пользователя в виде "get and set". Для этого объект **Document** предоставляет методы **getSelectedDataAsync** и **setSelectedDataAsync**.

Примеры кода, демонстрирующие выполнение задач с выделенными фрагментами, см. в статье [Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Работа с привязками с помощью объектов Bindings и Binding


Доступ к данным на основе привязок позволяет надстройкам области задач и контентным надстройкам получать единообразный доступ к определенной области документа или электронной таблицы через идентификатор, связанный с привязкой. Сначала надстройка должна создать привязку с помощью вызова одного из методов, связывающих часть документа с уникальным идентификатором: [addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) или [addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync). После настройки привязки надстройка может использовать предоставленный идентификатор для доступа к данным, содержащимся в связанном регионе документа или электронной таблицы. Создание привязок предоставляет указанные ниже возможности.


- Разрешает доступ к общим структурам данных в поддерживаемых приложениях Office, таким как: таблицы, диапазоны или текст (связанная последовательность знаков).
    
- Позволяет производить операции чтения или записи без необходимости выделения пользователем фрагмента.
    
- Устанавливает отношение между надстройкой и данными в документе. Привязки сохраняются в документе и могут использоваться позже.
    
Установка привязки также позволяет подписываться на данные и выбирать изменения событий, относящиеся к конкретной области документа или электронной таблицы. Это означает, что надстройка уведомляется только об изменениях, происходящих внутри данной конкретной области, в отличие от изменений, затрагивающих в целом весь документ или электронную таблицу.

Объект [Bindings](http://dev.office.com/reference/add-ins/shared/bindings.bindings) предоставляет метод [getAllAsync](http://dev.office.com/reference/add-ins/shared/bindings.getallasync), который обеспечивает доступ к набору всех привязок, установленных в этом документе или листе. Доступ к отдельной привязке можно получить по ее идентификатору с помощью методов [Bindings.getBindingByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) или [Office.select](http://dev.office.com/reference/add-ins/shared/office.select). Можно создать новые привязки, а также удалить существующие, используя один из перечисленных ниже методов объекта **Bindings**: [addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync), [addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), [addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) или [releaseByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync).

Имеется три различных вида привязки, которые определяются с помощью параметра  _bindingType_ при создании привязки с помощью методов **addFromSelectionAsync**, **addFromPromptAsync** или **addFromNamedItemAsync**:



|**Тип привязки**|**Описание**|**Поддержка ведущего приложения**|
|:-----|:-----|:-----|
|Привязка текста|Выполняет привязку к области документа, которая может быть представлена как текст.|В Word поддерживается большинство связанных выделений, тогда как в Excel для привязки текста можно использовать только выделения отдельных ячеек. Excel поддерживает только обычный текст, а Word — три формата: обычный текст, HTML и Open XML для Office.|
|Привязка матрицы|Выполняет привязку к фиксированной области документа, содержащей табличные данные без заголовков. Данные в привязке матрицы записываются или считываются как двумерный  **Array**, который в JavaScript реализован как массив массивов. Например, две строки значений **string** в двух столбцах можно записать или прочитать как ` [['a', 'b'], ['c', 'd']]`, а один столбец, состоящий из трех строк, — как `[['a'], ['b'], ['c']]`.|В Excel для установки матричной привязки может использоваться любое связанное выделение ячеек. В Word матричная привязка поддерживается только таблицами.|
|Привязка таблицы|Выполняет привязку к области документа, содержащей таблицу с заголовками. Данные в привязке таблицы записываются или считываются как объект [TableData](http://dev.office.com/reference/add-ins/shared/tabledata). Объект **TableData** предоставляет данные с помощью свойств **headers** и **rows**.|Любая таблица Excel или Word может быть основой для табличной привязки. После создания табличной привязки каждая новая строка или столбец, добавляемые пользователем в таблицу, автоматически включаются в привязку. |
После создания привязки с помощью одного из трех методов add объекта **Bindings** можно работать с данными и свойствами привязки с помощью методов соответствующего объекта: [MatrixBinding](http://dev.office.com/reference/add-ins/shared/binding.matrixbinding), [TableBinding](http://dev.office.com/reference/add-ins/shared/binding.tablebinding) или [TextBinding](http://dev.office.com/reference/add-ins/shared/binding.textbinding). Все три эти объекта наследуют методы [getDataAsync](http://dev.office.com/reference/add-ins/shared/binding.getdataasync) и [setDataAsync](http://dev.office.com/reference/add-ins/shared/binding.setdataasync) объекта **Binding**, позволяющие взаимодействовать с привязанными данными.

Примеры кода, демонстрирующие выполнение задач с привязками, см. в статье [Привязка к областям в документе или электронной таблице](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>Работа с настраиваемыми частями XML с помощью объектов CustomXmlParts и CustomXmlPart


 **Область применения:** надстройки области задач Word

Объекты [CustomXmlParts](http://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) и [CustomXmlPart](http://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) интерфейса API предоставляют доступ к настраиваемым частям XML в документах Word, которые позволяют работать с содержимым документа на основе XML. Примеры работы с объектами **CustomXmlParts** и **CustomXmlPart** см. в примере кода [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts).


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>Работа с целым документом с помощью метода getFileAsync


 **Область применения:** надстройки области задач Word и PowerPoint

Метод [Document.getFileAsync](http://dev.office.com/reference/add-ins/shared/document.getfileasync) и члены объектов [File](http://dev.office.com/reference/add-ins/shared/file) и [Slice](http://dev.office.com/reference/add-ins/shared/slice) предоставляют функциональность получения целого файла документа Word и PowerPoint в виде порций (блоков) до 4 МБ единовременно. Дополнительные сведения см. в разделе [Получение всего содержимого файла из документа в надстройке](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md).


## <a name="mailbox-object"></a>Объект Mailbox


 **Область применения:** надстройки Outlook

Надстройки Outlook, в основном, используют набор API, предоставляемый через объект [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox). Чтобы получить объекты и члены специально для использования в надстройках Outlook, такие как объект [Item](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item), используйте свойство [mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) объекта **Context** для получения доступа к объекту **Mailbox**, как показано в следующей строке кода.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Кроме того, надстройки Outlook могут использовать следующие объекты:


-  Объект **Office** для инициализации.
    
-  Объект **Context** для получения доступа к контенту и отображения языковых свойств.
    
-  Объект **RoamingSettings** для сохранения в почтовом ящике пользователя, в котором установлено приложение, пользовательских свойств, относящихся к надстройке Outlook.
    
Сведения об использовании JavaScript в надстройках Outlook см. в статьях [Надстройки Outlook](../outlook/outlook-add-ins.md) и [Обзор архитектуры и функций надстроек Outlook](../outlook/overview.md).


## <a name="api-support-matrix"></a>Матрица поддержки API


В данной таблице представлены API и функции, поддерживаемые всеми типами надстроек (контентными, области задач и Outlook), и приложения Office, которые могут их размещать, когда вы указываете [ведущие приложения Office, поддерживаемые вашей надстройкой](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) с помощью [схемы манифестов надстроек версии 1.1 и функций, поддерживаемых API JavaScript для Office версии 1.1](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**Имя узла**|База данных|Книга|Почтовый ящик|Презентация|Документ|Project|
||**Поддерживаемые** **ведущие приложения**|Веб-приложения Access|ExcelExcel Online|OutlookOutlook Web AppOutlook Web App для устройств|PowerPointPower Point Online|Word|Project|
|**Поддерживаемые типы надстроек**|Содержимое|Да|Да||Да|||
||Область задач||Да||Да|Да|Да|
||Outlook|||Да||||
|**Поддерживаемые функции API**|Чтение и запись текста||Да||Да|Да|Да (только чтение)|
||Матрица чтения и записи||Да|||Да||
||Таблица чтения и записи||Да|||Да||
||Чтение и запись HTML|||||Да||
||Чтение и запись Office Open XML|||||Да||
||Чтение свойств задач, ресурсов, представлений и полей||||||Да|
||События изменения выделенного фрагмента||Да|||Да||
||Загрузка всего документа||||Да|Да||
||Привязки и их события|Да (только полные и частичные привязки таблиц)|Да|||Да||
||Чтение и запись. Настраиваемые части XML|||||Да||
||Сохранение данных состояния надстройки(параметры)|Да (для каждой ведущей надстройки)|Да (для каждого документа)|Да (для каждого почтового ящика)|Да (для каждого документа)|Да (для каждого документа)||
||События изменения параметров|Да|Да||Да|Да||
||События получения активного режима просмотра и изменения представления||||Да|||
||Переход к расположениям в документе||Да||Да|Да||
||Контекстная активация с использованием правил и регулярных выражений|||Да||||
||Чтение свойств элемента|||Да||||
||Чтение профиля пользователя|||Да||||
||Получение вложений|||Да||||
||Получение токена удостоверения пользователя|||Да||||
||Вызов веб-служб Exchange|||Да||||
