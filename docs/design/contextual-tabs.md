---
title: Создание настраиваемой контекстной вкладки в надстройки Office
description: Узнайте, как добавлять настраиваемые контекстные вкладки в надстройку Office.
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: 3939e3338c734e1d6400dc261b59e35de63e5779
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789137"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Создание настраиваемой контекстной вкладки в надстройки Office (предварительная версия)

Контекстная вкладка — это скрытая вкладка на ленте Office, которая отображается в строке вкладки, когда в документе Office происходит определенное событие. Например, **вкладка "Конструктор таблицы",** которая отображается на ленте Excel при выборе таблицы. Вы можете включить настраиваемые контекстные вкладки в надстройку Office и указать, когда они будут видимыми или скрытыми, создав обработчики событий, которые меняют видимость. (Однако настраиваемые контекстные вкладки не реагируют на изменения фокуса.)

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

> [!IMPORTANT]
> Настраиваемые контекстные вкладки находятся в предварительной версии. Поэкспериментируйте с ними в среде разработки или тестирования, но не добавляйте их в производственную надстройки.
>
> Настраиваемые контекстные вкладки в настоящее время поддерживаются только в Excel и только на этих платформах и сборках:
>
> - Excel для Windows (только Microsoft 365, а не бессрочная лицензия): версия 2011 (сборка 13426.20274). Возможно, ваша подписка на Microsoft 365 должна быть на канале [Current Channel (предварительная версия),](https://insider.office.com/join/windows) который ранее назывался "Monthly Channel (Targeted)" или "Insider Slow".

> [!NOTE]
> Настраиваемые контекстные вкладки работают только на платформах, которые поддерживают следующие наборы требований. Дополнительные информацию о наборах требований и работе с ними см. в подразделе "Указание приложений [Office и требований к API".](../develop/specify-office-hosts-and-api-requirements.md)
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>Поведение настраиваемой контекстной вкладки

Пользовательский интерфейс настраиваемой контекстной вкладки следует шаблону встроенных контекстных вкладок Office. Ниже основных принципов размещения настраиваемой контекстной вкладки:

- Когда пользовательская контекстная вкладка отображается, она отображается в правой части ленты.
- Если одна или несколько встроенных контекстных вкладок и одна или несколько настраиваемые контекстные вкладки из надстроек видны одновременно, настраиваемые контекстные вкладки всегда находятся справа от всех встроенных контекстных вкладок.
- Если надстройка имеет несколько контекстных вкладок и существуют контексты, в которых отображается несколько, они отображаются в том порядке, в котором они определены в надстройке. (Направление в том же направлении, что и язык Office, то есть направление слева направо на языках слева направо, а направление справа налево — на языках справа налево.) Подробные [сведения о том,](#define-the-groups-and-controls-that-appear-on-the-tab) как их определить, см. в поднаборе "Определение групп и элементов управления, которые отображаются на вкладке".
- Если несколько надстроек имеет контекстную вкладку, которая отображается в определенном контексте, они отображаются в том порядке, в котором были запущены надстройки.
- Настраиваемые *контекстные* вкладки, в отличие от настраиваемой основной вкладки, не добавляются окончательно на ленту приложения Office. Они присутствуют только в документах Office, в которых работает надстройка.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Основные действия по добавлению контекстной вкладки в надстройку

Далее приводится основной этап добавления настраиваемой контекстной вкладки в надстройку.

1. Настройте надстройку для использования общей времени работы.
1. Определите вкладку, группы и элементы управления, которые отображаются на ней.
1. Зарегистрируйте контекстную вкладку в Office.
1. Укажите условия, в которые вкладка будет видна.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общей времени работы

Чтобы добавить настраиваемые контекстные вкладки, надстройка будет использовать общую времени работы. Дополнительные сведения см. в настройках [надстройки для использования общей времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Определение групп и элементов управления, которые отображаются на вкладке

В отличие от настраиваемой основной вкладки, которые определены с помощью XML в манифесте, настраиваемые контекстные вкладки определяются во время работы с BLOB JSON. Код разбрасирует большой объект в объект JavaScript, а затем передает объект [методу Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) Настраиваемые контекстные вкладки присутствуют только в документах, в которых в настоящее время работает надстройка. Это отличается от настраиваемой основной вкладки, которые добавляются на ленту приложения Office при установке надстройки и остаются в настоящем при открытом другом документе. Кроме того, `requestCreateControls` метод можно запустить только один раз в сеансе надстройки. Если он будет вызван повторно, будет выброшена ошибка.

> [!NOTE]
> Структура свойств и подэлементов BLOB-объекта JSON (и имен ключей) приблизительно параллельна структуре элемента [CustomTab](../reference/manifest/customtab.md) и его потомков в XML манифеста.

Пошаговое создание примера контекстных вкладок JSON. (Полная схема контекстной вкладки JSON находится [вdynamic-ribbon.schema.js.](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json) Эта ссылка может не работать в начале периода предварительного просмотра для контекстных вкладок. Если ссылка не работает, вы можете найти последний черновик схемы на черновике dynamic-ribbon.schema.js[на](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) Если вы работаете в Visual Studio Code, этот файл можно использовать для получения IntelliSense проверки JSON. Дополнительные сведения см. в редактировании [JSON с помощью Visual Studio Code — схемы и параметры JSON.](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)


1. Начните с создания строки JSON с двумя свойствами массива с именем `actions` и `tabs` . Массив — это спецификация всех функций, которые можно выполнять с помощью элементов `actions` управления на контекстной вкладке. Массив определяет одну или несколько контекстных вкладок `tabs` до *10.*

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Этот простой пример контекстной вкладки будет иметь только одну кнопку и, следовательно, только одно действие. Добавьте следующий как единственный член `actions` массива. Обратите внимание на эту разметку:

    - Свойства `id` являются `type` обязательными.
    - Значением `type` может быть ExecuteFunction или ShowTaskpane.
    - Свойство `functionName` используется только в том случае, если `type` значением является `ExecuteFunction` . Это имя функции, определенной в FunctionFile. Дополнительные сведения о FunctionFile см. в основных понятиях для команд [надстройки.](add-in-commands.md)
    - На более позднем этапе вы соберем это действие с кнопкой на контекстной вкладке.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Добавьте следующий как единственный член `tabs` массива. Обратите внимание на эту разметку:

    - Свойство `id` является обязательным. Используйте краткий описательный ИД, уникальный для всех контекстных вкладок в надстройке.
    - Свойство `label` является обязательным. Это пользовательская строка, которая служит меткой контекстной вкладки.
    - Свойство `groups` является обязательным. Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен иметь по крайней мере один член *и не более 20*. (Кроме того, существуют ограничения на количество элементов управления, которые можно использовать на настраиваемой контекстной вкладке, а также количество групп. Дополнительные сведения см. в следующем шаге.)

    > [!NOTE]
    > Кроме того, у объекта tab может быть необязательное свойство, которое указывает, отображается ли вкладка сразу после начала `visible` надстройки. Так как контекстные вкладки обычно скрыты до тех пор, пока событие пользователя не активирует их видимость (например, пользователь выбирает объект того или иного типа в документе), свойство по умолчанию имеет значение, когда его `visible` `false` нет. В более позднем разделе мы покажем, как настроить свойство `true` в ответ на событие.

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. В простом примере контекстная вкладка имеет только одну группу. Добавьте следующий как единственный член `groups` массива. Обратите внимание на эту разметку:

    - Все свойства являются обязательной.
    - Свойство должно быть уникальным для всех групп на `id` вкладке. Используйте краткий и описательный ИД.
    - Это `label` пользовательская строка, которая будет служить меткой группы.
    - Значение свойства — это массив объектов, которые указывают значки, которые будут иметься в группе на ленте в зависимости от размера ленты и окна `icon` приложения Office.
    - Значение свойства — это массив объектов, которые указывают кнопки и меню `controls` в группе. В группе должно быть по крайней мере один и не более *6.*

    > [!IMPORTANT]
    > *Общее число элементов управления на всей вкладке не может быть больше 20.* Например, можно иметь 3 группы с по 6 элементов управления и четвертую группу с 2 элементами управления, но нельзя иметь 4 группы с по 6 элементов управления в каждой.  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. Каждая группа должна иметь значок размером не менее двух размеров: 32x32 пк и 80x80 пк. Кроме того, можно использовать значки размером 16x16 пк, 20x20 px, 24x24 px, 40x40 px, 48x48 px и 64x64 px. Office определяет, какой значок использовать в зависимости от размера ленты и окна приложения Office. Добавьте следующие объекты в массив значков. (Если размер окна и ленты достаточно велик для  появления хотя бы одного из элементов управления в группе, значок группы вообще не отображается. Например, просмотрите группу **стилей** на ленте Word при сжатии и расширении окна Word.) Обратите внимание на эту разметку:

    - Оба свойства являются обязательной.
    - Единица `size` измерения свойства — пиксели. Значки всегда квадратные, поэтому числом является и высота, и ширина.
    - Свойство `sourceLocation` указывает полный URL-адрес значка.

    > [!IMPORTANT]
    > Так же, как обычно необходимо изменить URL-адреса в манифесте надстройки при переходе от разработки к производственной (например, при изменении домена с localhost на contoso.com), необходимо также изменить URL-адреса в контекстных вкладок JSON.

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. В нашем простом примере группа имеет только одну кнопку. Добавьте следующий объект в качестве единственный член `controls` массива. Обратите внимание на эту разметку:

    - Все свойства, кроме `enabled` , являются обязательной.
    - `type` указывает тип управления. Значениями могут быть "Button", "Menu" или "MobileButton".
    - `id` может быть до 125 символов. 
    - `actionId` должен быть ИД действия, определенного в `actions` массиве. (См. шаг 1 этого раздела.)
    - `label` — это пользовательская строка, которая служит в качестве подписи кнопки.
    - `superTip` представляет собой форматную форму подсказки. Необходимы `title` и `description` свойства, и свойства.
    - `icon` указывает значки для кнопки. Здесь также применимы предыдущие замечания о значке группы.
    - `enabled` (необязательно) указывает, включена ли кнопка при отжатии контекстной вкладки. Значение по умолчанию , если нет `true` . 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
Ниже приводится полный пример BLOB-примера JSON:

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Регистрация контекстной вкладки в Office с помощью requestCreateControls

Контекстная вкладка регистрируется в Office путем вызова метода [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Обычно это делается либо в функции, назначенной методу, либо `Office.initialize` с помощью `Office.onReady` этого метода. Подробнее об этих методах и инициализации надстройки см. в инициализации [надстройки Office.](../develop/initialize-add-in.md) Однако вы можете вызвать метод в любое время после инициализации.

> [!IMPORTANT]
> Метод `requestCreateControls` может быть вызван только один раз в заданном сеансе надстройки. Если она будет вызвана повторно, будет выброшена ошибка.

Ниже приведен пример. Обратите внимание, что перед тем как передать строку JSON в функцию JavaScript, ее необходимо преобразовать в объект JavaScript с помощью `JSON.parse` метода.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Укажите контексты, когда вкладка будет видна с помощью requestUpdate

Как правило, настраиваемая контекстная вкладка должна отображаться, когда событие, инициированное пользователем, изменяет контекст надстройки. Рассмотрим сценарий, в котором вкладка должна быть видна, когда и только когда активируется диаграмма (на стандартной таблице книги Excel).

Начните с назначения обработчиков. Обычно это делается в методе, как в следующем примере, который назначает обработчики (созданные на более позднем этапе) всем диаграммам на `Office.onReady` `onActivated` этом `onDeactivated` графике и событиям.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Затем определите обработчики. Ниже приводится простой пример ошибки `showDataTab` [HostRestartNeeded,](#handling-the-hostrestartneeded-error) но более надежную версию функции см. далее в этой статье. Вот что нужно знать об этом коде:

- Office определяет время обновления состояния ленты. Метод  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) очереди запрос на обновление. Метод разрешит объект сразу после того, как он задюет запрос в очередь, а не при `Promise` обновлении ленты.
- Параметром метода является объект `requestUpdate` [RibbonUpdaterData,](/javascript/api/office/office.ribbonupdaterdata) который (1) указывает вкладку по ее ИД точно так же, как указано в *JSON* и (2) определяет видимость вкладки.
- Если имеется несколько настраиваемой контекстной вкладки, которая должна быть видна в одном контексте, в массив просто добавляются дополнительные объекты `tabs` вкладок.

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

Обработок для скрытие вкладки почти идентичен, за исключением того, что он возвращает `visible` `false` свойство.

Библиотека JavaScript для Office также предоставляет несколько интерфейсов (типов), упрощая создание `RibbonUpdateData` объекта. Ниже приводится функция `showDataTab` в TypeScript, которая использует эти типы.

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Одновременное переключеть видимость вкладки и состояние включения кнопки

Этот метод также используется для включения или отключения состояния настраиваемой кнопки на настраиваемой контекстной вкладке или в пользовательской `requestUpdate` основной вкладке. Дополнительные сведения см. в подстройке "Включить и отключить [команды надстройки".](disable-add-in-commands.md) Возможны сценарии, в которых одновременно необходимо изменить видимость вкладки и состояние кнопки. Это можно сделать одним вызовом `requestUpdate` . Ниже приводится пример, в котором кнопка на основной вкладке включена одновременно с видимой контекстной вкладками.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

В следующем примере включенная кнопка находится на той же контекстной вкладке, которая отображается.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="localizing-the-json-blob"></a>Локализация BLOB JSON

Передаваемый BLOB-проект JSON не локализуется так же, как локализована разметка манифеста для настраиваемой основной вкладки (что описано в локализации control из `requestCreateControls` манифеста). [](../develop/localization.md#control-localization-from-the-manifest) Вместо этого локализация должна происходить во время работы с использованием отдельных BLOB-ок JSON для каждого из региональных стандартов. Мы рекомендуем использовать заявление, которое тестирует `switch` [свойство Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage) Ниже приведен пример.

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

Затем код вызывает функцию, чтобы получить локализованный BLOB-код, который передается в, как в `requestCreateControls` следующем примере:

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="handling-the-hostrestartneeded-error"></a>Обработка ошибки HostRestartNeeded

В некоторых случаях Office не может обновить ленту и возвращает ошибку. Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть. Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`. Ниже приведен пример обработки этой ошибки. В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
