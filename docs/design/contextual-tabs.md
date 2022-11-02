---
title: Создание пользовательских контекстных вкладок в надстройках Office
description: Узнайте, как добавить пользовательские контекстные вкладки в надстройку Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f43f6ec0a6ef3faef4c5e50d5da6d124124fe92
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810234"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Создание пользовательских контекстных вкладок в надстройках Office

Контекстная вкладка — это скрытый элемент управления tab на ленте Office, который отображается в строке вкладки при возникновении указанного события в документе Office. Например, вкладка **"Конструктор таблиц** ", которая появляется на ленте Excel при выборе таблицы. Вы включаете пользовательские контекстные вкладки в надстройку Office и указываете, когда они видны или скрыты, создавая обработчики событий, которые изменяют видимость. (Однако пользовательские контекстные вкладки не реагируют на изменения фокуса.)

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

> [!IMPORTANT]
> Пользовательские контекстные вкладки в настоящее время поддерживаются только в Excel и только на этих платформах и сборках.
>
> - Excel в Windows: версия 2102 (сборка 13801.20294) или более поздняя.
> - Excel для Mac: версия 16.53.806.0 или более поздняя.
> - Excel в Интернете

> [!NOTE]
> Пользовательские контекстные вкладки работают только на платформах, поддерживающих следующие наборы требований. Дополнительные сведения о наборах требований и способах работы с ними см [. в разделе Указание приложений Office и требований к API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
> - [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
>
> Вы можете использовать проверки среды выполнения в коде, чтобы проверить, поддерживает ли сочетание узла и платформы пользователя эти наборы требований, как описано в разделе [Проверка среды выполнения для поддержки методов и наборов требований](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (Метод указания наборов требований в манифесте, который также описан в этой статье, в настоящее время не работает для RibbonApi 1.2.) Кроме того, можно [реализовать альтернативный интерфейс пользовательского интерфейса, если пользовательские контекстные вкладки не поддерживаются](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Поведение пользовательских контекстных вкладок

Пользовательский интерфейс для пользовательских контекстных вкладок соответствует шаблону встроенных контекстных вкладок Office. Ниже приведены основные принципы размещения пользовательских контекстных вкладок.

- Когда отображается пользовательская контекстная вкладка, она отображается на правом конце ленты.
- Если одна или несколько встроенных контекстных вкладок и одна или несколько настраиваемых контекстных вкладок из надстроек видны одновременно, настраиваемые контекстные вкладки всегда находятся справа от всех встроенных контекстных вкладок.
- Если надстройка содержит несколько контекстных вкладок и есть контексты, в которых отображается несколько, они отображаются в том порядке, в котором они определены в надстройке. (Направление совпадает с направлением языка Office. То есть слева направо в языках слева направо, но справа налево в языках справа налево.) Дополнительные сведения о том, как вы их [определяете, см. в статье Определение групп и элементов управления, отображаемых на вкладке](#define-the-groups-and-controls-that-appear-on-the-tab) .
- Если несколько надстроек имеют контекстные вкладки, видимые в определенном контексте, они отображаются в том порядке, в котором были запущены надстройки.
- *Пользовательские контекстные* вкладки, в отличие от пользовательских основных вкладок, не добавляются на ленту приложения Office. Они присутствуют только в документах Office, в которых выполняется надстройка.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Основные действия по включению контекстной вкладки в надстройку

Ниже приведены основные шаги по включению пользовательской контекстной вкладки в надстройку.

1. Настройте надстройку для использования общей среды выполнения.
1. Определите вкладку, а также группы и элементы управления, которые на ней отображаются.
1. Зарегистрируйте контекстную вкладку в Office.
1. Укажите обстоятельства, когда вкладка будет видна.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общей среды выполнения

Для добавления пользовательских контекстных вкладок надстройка должна использовать [общую среду выполнения](../testing/runtimes.md#shared-runtime). Дополнительные сведения см. [в разделе Настройка надстройки для использования общей среды выполнения](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Определение групп и элементов управления, отображаемых на вкладке

В отличие от пользовательских основных вкладок, которые определяются с помощью XML в манифесте, пользовательские контекстные вкладки определяются во время выполнения с помощью большого двоичного объекта JSON. Код анализирует большой двоичный объект в объект JavaScript, а затем передает объект в метод [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . Пользовательские контекстные вкладки присутствуют только в документах, в которых сейчас выполняется надстройка. Это отличается от пользовательских основных вкладок, которые добавляются на ленту приложения Office при установке надстройки и сохраняются при открытии другого документа. Кроме того, `requestCreateControls` метод может выполняться только один раз в сеансе надстройки. При повторном вызове возникает ошибка.

> [!NOTE]
> Структура свойств и вложенных свойств большого двоичного объекта JSON (и имен ключей) примерно аналогична структуре элемента [CustomTab](/javascript/api/manifest/customtab) и его потомков в XML-коде манифеста.

Мы создадим пример большого двоичного объекта JSON с контекстными вкладками. Полная схема для JSON контекстной вкладки находится в [файле dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Если вы работаете в Visual Studio Code, этот файл можно использовать для получения IntelliSense и проверки JSON. Дополнительные сведения см. в разделе [Изменение JSON с помощью Visual Studio Code — схемы и параметры JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Начните с создания строки JSON с двумя свойствами массива с именами `actions` и `tabs`. Массив `actions` — это спецификация всех функций, которые могут выполняться элементами управления на контекстной вкладке. Массив `tabs` определяет одну или несколько контекстных вкладок, *максимум до 20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Этот простой пример контекстной вкладки будет содержать только одну кнопку и, следовательно, только одно действие. Добавьте следующий элемент в качестве единственного члена массива `actions` . Обратите внимание на эту разметку:

    - `id` Свойства и `type` являются обязательными.
    - Значение `type` может быть либо "ExecuteFunction", либо "ShowTaskpane".
    - Свойство `functionName` используется только в том случае, если значение равно `type` `ExecuteFunction`. Это имя функции, определенной в FunctionFile. Дополнительные сведения о FunctionFile см. в разделе [Основные понятия для команд надстроек](add-in-commands.md).
    - На следующем шаге это действие будет сопоставлено с кнопкой на контекстной вкладке.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. Добавьте следующий элемент в качестве единственного члена массива `tabs` . Обратите внимание на эту разметку:

    - Свойство `id` является обязательным. Используйте краткий описательный идентификатор, уникальный среди всех контекстных вкладок в надстройке.
    - Свойство `label` является обязательным. Это удобная строка, которая служит меткой контекстной вкладки.
    - Свойство `groups` является обязательным. Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен содержать по крайней мере один член *и не более 20*. (Существуют также ограничения на количество элементов управления, которые можно использовать на пользовательской контекстной вкладке, что также ограничивает количество имеющихся групп. Дополнительные сведения см. в следующем шаге.)

    > [!NOTE]
    > Объект tab также может иметь необязательное `visible` свойство, указывающее, отображается ли вкладка сразу при запуске надстройки. Так как контекстные вкладки обычно скрыты до тех пор, пока событие пользователя не активирует видимость (например, пользователь выбирает сущность определенного типа в документе), `visible` свойство по умолчанию имеет значение , `false` если его нет. В следующем разделе мы покажем, как задать свойству значение `true` в ответ на событие.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. В простом текущем примере контекстная вкладка содержит только одну группу. Добавьте следующий элемент в качестве единственного члена массива `groups` . Обратите внимание на эту разметку:

    - Все свойства являются обязательными.
    - Свойство `id` должно быть уникальным среди всех групп в манифесте. Используйте краткий описательный идентификатор, который содержит до 125 символов.
    - — `label` это удобная строка, которая служит меткой группы.
    - Значение `icon` свойства — это массив объектов, указывающий значки, которые группа будет иметь на ленте в зависимости от размера ленты и окна приложения Office.
    - Значение `controls` свойства представляет собой массив объектов, указывающих кнопки и меню в группе. Должен быть хотя бы один.

    > [!IMPORTANT]
    > *Общее количество элементов управления на всей вкладке может быть не более 20.* Например, вы можете иметь 3 группы с 6 элементами управления каждый и четвертую группу с 2 элементами управления, но вы не можете иметь 4 группы с 6 элементами управления.  

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

1. Каждая группа должна иметь значок не менее двух размеров: 32x32 пикселей и 80x80 пикселей. При необходимости можно также иметь значки размером 16x16 пикселей, 20x20 пикселей, 24x24 пикселей, 40x40 пикселей, 48x48 пикселей и 64x64 пикселей. Office определяет, какой значок следует использовать, в зависимости от размера ленты и окна приложения Office. Добавьте следующие объекты в массив значков. (Если размер окна и ленты достаточно велик для отображения хотя бы одного из *элементов управления* в группе, значок группы вообще не отображается. Например, просмотрите группу **Стили** на ленте Word при сжатии и развертывании окна Word.) Обратите внимание на эту разметку:

    - Оба свойства являются обязательными.
    - Единица `size` измерения свойства — пиксели. Значки всегда квадратные, поэтому число равно высоте и ширине.
    - Свойство `sourceLocation` указывает полный URL-адрес значка.

    > [!IMPORTANT]
    > Как правило, при переходе от разработки к рабочей среде необходимо изменить URL-адреса в манифесте надстройки (например, при изменении домена с localhost на contoso.com), необходимо также изменить URL-адреса в json контекстных вкладках.

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

1. В нашем простом текущем примере группа имеет только одну кнопку. Добавьте следующий объект в качестве единственного члена массива `controls` . Обратите внимание на эту разметку:

    - Все свойства, кроме `enabled`, являются обязательными.
    - `type` указывает тип элемента управления. Значения могут быть "Button", "Menu" или "MobileButton".
    - `id` Может содержать до 125 символов.
    - `actionId` должен быть идентификатором действия, определенного в массиве `actions` . (См. шаг 1 этого раздела.)
    - `label` — это удобная для пользователя строка, которая служит меткой кнопки.
    - `superTip` представляет богатую форму подсказки. `title` Свойства и являются `description` обязательными.
    - `icon` указывает значки для кнопки. Здесь также применяются предыдущие замечания о значке группы.
    - `enabled` (необязательно) указывает, включена ли кнопка при запуске контекстной вкладки. Значение по умолчанию, если оно отсутствует, — .`true`

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

Ниже приведен полный пример большого двоичного объекта JSON.

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
      "label": "Contoso Data",
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

Контекстная вкладка регистрируется в Office путем вызова метода [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . Обычно это выполняется в функции, назначенной `Office.initialize` или с `Office.onReady` помощью функции . Дополнительные сведения об этих функциях и инициализации надстройки см. в статье [Инициализация надстройки Office](../develop/initialize-add-in.md). Однако метод можно вызвать в любое время после инициализации.

> [!IMPORTANT]
> Метод `requestCreateControls` может вызываться только один раз в заданном сеансе надстройки. При повторном вызове возникает ошибка.

Ниже приведен пример. Обратите внимание, что перед передачей в функцию JavaScript строку JSON необходимо преобразовать в объект JavaScript с `JSON.parse` помощью метода .

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Укажите контексты, когда вкладка будет отображаться с помощью requestUpdate

Как правило, пользовательская контекстная вкладка должна отображаться при изменении контекста надстройки событием, инициированным пользователем. Рассмотрим сценарий, в котором вкладка должна отображаться только при активации диаграммы (на листе книги Excel по умолчанию).

Начните с назначения обработчиков. Обычно это делается в `Office.onReady` функции, как в следующем примере, которая назначает обработчики (созданные на более позднем шаге) `onActivated` событиям и `onDeactivated` всех диаграмм на листе.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Затем определите обработчики. Ниже приведен простой пример , но более надежная `showDataTab`версия функции см. в статье [Обработка ошибки HostRestartNeeded](#handle-the-hostrestartneeded-error) далее в этой статье. Вот что нужно знать об этом коде:

- Office определяет время обновления состояния ленты. Метод  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) помещает запрос на обновление в очередь. Метод разрешает объект, `Promise` как только он помещает запрос в очередь, а не при фактическом обновлении ленты.
- Параметром метода `requestUpdate` является объект [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) , который (1) указывает вкладку по идентификатору *точно так, как указано в JSON* , и (2) указывает видимость вкладки.
- Если у вас есть несколько пользовательских контекстных вкладок, которые должны быть видны в одном контексте, просто добавьте дополнительные объекты табуляции в `tabs` массив.

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

Обработчик для скрытия вкладки почти идентичен, за исключением того, что он задает свойству `visible` значение `false`.

Библиотека JavaScript для Office также предоставляет несколько интерфейсов (типов), упрощая создание`RibbonUpdateData` объекта. Ниже приведена `showDataTab` функция в TypeScript, и она использует эти типы.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Одновременное включение видимости вкладки и состояние включенной кнопки

Метод `requestUpdate` также используется для переключения состояния включенной или отключенной пользовательской кнопки на настраиваемой контекстной вкладке или на настраиваемой вкладке core. Дополнительные сведения об этом см. [в разделе Включение и отключение команд надстроек](disable-add-in-commands.md). Могут возникать сценарии, в которых требуется одновременно изменить видимость вкладки и состояние включенной кнопки. Это можно сделать с помощью одного вызова `requestUpdate`. Ниже приведен пример, в котором кнопка на основной вкладке включается в то же время, когда контекстная вкладка становится видимой.

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
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

В следующем примере включенная кнопка находится на той же контекстной вкладке, которая становится видимой.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## <a name="open-a-task-pane-from-contextual-tabs"></a>Открытие области задач из контекстных вкладок

Чтобы открыть область задач с помощью кнопки на настраиваемой контекстной вкладке, создайте действие в ФОРМАТЕ JSON с `type` .`ShowTaskpane` Затем определите кнопку со свойством `actionId` , присвоенным свойству `id` для действия . Откроется область задач по умолчанию, указанная элементом **\<Runtime\>** манифеста.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

Чтобы открыть любую область задач, которая не является областью задач по умолчанию, укажите `sourceLocation` свойство в определении действия. В следующем примере вторая область задач открывается с другой кнопки.

> [!IMPORTANT]
>
> - `sourceLocation` Если для действия задано значение , область задач *не* использует общую среду выполнения. Он выполняется в новой отдельной среде выполнения.
> - Не более одной области задач могут использовать общую среду выполнения, поэтому не более одного действия типа `ShowTaskpane` могут опустить `sourceLocation` свойство .

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a>Локализация текста JSON

Большой двоичный объект JSON, передаваемый `requestCreateControls` в, не локализуется так же, как и разметка манифеста для пользовательских основных вкладок (что описано в разделе [Локализация элемента управления из манифеста](../develop/localization.md#control-localization-from-the-manifest)). Вместо этого локализация должна выполняться во время выполнения с использованием отдельных больших двоичных объектов JSON для каждого языкового стандарта. Рекомендуется использовать `switch` оператор, который проверяет свойство [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) . Ниже приведен пример.

```javascript
function GetContextualTabsJsonSupportedLocale () {
    const displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
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
                          "label": "Contoso Données",
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

Затем код вызывает функцию для получения локализованного большого двоичного объекта, переданного `requestCreateControls`в , как показано в следующем примере.

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Рекомендации по пользовательским контекстным вкладкам

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Реализация альтернативного интерфейса пользовательского интерфейса, если пользовательские контекстные вкладки не поддерживаются

Некоторые сочетания платформы, приложения Office и сборки Office не поддерживают `requestCreateControls`. Ваша надстройка должна быть разработана таким образом, чтобы обеспечить альтернативный интерфейс для пользователей, которые выполняют надстройку в одном из этих сочетаний. В следующих разделах описаны два способа предоставления резервного взаимодействия.

#### <a name="use-noncontextual-tabs-or-controls"></a>Использование неконтекстуальных вкладок или элементов управления

Существует элемент манифеста [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi), предназначенный для создания резервного интерфейса в надстройке, которая реализует пользовательские контекстные вкладки, когда надстройка выполняется в приложении или на платформе, которая не поддерживает настраиваемые контекстные вкладки.

Самая простая стратегия использования этого элемента заключается в определении настраиваемой вкладки core (т. е. *неконтекстовой* настраиваемой вкладки) в манифесте, которая дублирует настройки ленты настраиваемых контекстных вкладок в надстройке. Но вы добавляете `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` в качестве первого дочернего элемента повторяющихся элементов [Group](/javascript/api/manifest/group), [Control](/javascript/api/manifest/control) и menu **\<Item\>** на пользовательских основных вкладках. Результатом этого является следующее:

- Если надстройка выполняется в приложении и на платформе, которые поддерживают пользовательские контекстные вкладки, пользовательские основные группы и элементы управления не будут отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет создана, когда надстройка `requestCreateControls` вызывает метод .
- Если надстройка выполняется в приложении или на платформе, *которая не* поддерживает `requestCreateControls`, элементы будут отображаться на вкладке настраиваемого ядра.

Ниже приведен пример. Обратите внимание, что "MyButton" будет отображаться на настраиваемой вкладке core только в том случае, если пользовательские контекстные вкладки не поддерживаются. Но родительская группа и настраиваемая базовая вкладка будут отображаться независимо от того, поддерживаются ли пользовательские контекстные вкладки.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="Contoso.MyButton1">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

Дополнительные примеры см. в разделе [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi).

Если родительская группа или меню помечается параметром `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, он не отображается, а вся ее дочерняя разметка игнорируется, если пользовательские контекстные вкладки не поддерживаются. Таким образом, не имеет значения, содержит **\<OverriddenByRibbonApi\>** ли какой-либо из этих дочерних элементов элемент или его значение. Это связано с тем, что если элемент меню или элемент управления должен быть виден во всех контекстах, то он не только не должен быть помечен с `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`помощью , но *и его предки меню и группы также не должны быть помечены таким образом*.

> [!IMPORTANT]
> Не помечайте *все* дочерние элементы группы или меню с помощью `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`. Это бессмысленно, если родительский элемент помечается по `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` причинам, указанным в предыдущем абзаце. Кроме того, если оставить **\<OverriddenByRibbonApi\>** значение для родительского элемента (или задать для него значение `false`), родительский элемент будет отображаться независимо от того, поддерживаются ли пользовательские контекстные вкладки, но при их поддержке он будет пустым. Таким образом, если при поддержке пользовательских контекстных вкладок не должны отображаться все дочерние элементы, пометьте родительский элемент .`<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Использование API- интерфейсов, отображающих или скрывающих область задач в указанных контекстах

В качестве альтернативы **\<OverriddenByRibbonApi\>** надстройка может определить область задач с элементами управления пользовательского интерфейса, которые дублируют функциональные возможности элементов управления на пользовательской контекстной вкладке. Затем используйте методы [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) и [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) , чтобы отобразить область задач, когда контекстная вкладка была бы показана, если бы она была поддерживаема. Дополнительные сведения об использовании этих методов см. в статье [Отображение или скрытие области задач надстройки Office](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Обработка ошибки HostRestartNeeded

В некоторых случаях Office не может обновить ленту и возвращает ошибку. Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть. Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`. Код должен обрабатывать эту ошибку. Ниже приведен пример того, как это происходит. В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
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

## <a name="resources"></a>Ресурсы

- [Пример кода: создание пользовательских контекстных вкладок на ленте](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- Пример демонстрации контекстных вкладок сообщества

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]
