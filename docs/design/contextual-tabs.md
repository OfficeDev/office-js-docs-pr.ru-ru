---
title: Создание настраиваемых контекстных вкладок в надстройки Office
description: Узнайте, как добавлять настраиваемые контекстные вкладки в надстройку Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 23f6c64d1b3f0e95b8dcae6bc36563566acb8b3f
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958538"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Создание настраиваемых контекстных вкладок в надстройки Office

Контекстная вкладка — это скрытая вкладка на ленте Office, которая отображается в строке вкладки при возникновении указанного события в документе Office. Например, **вкладка "Конструктор таблиц"** , которая отображается на ленте Excel при выборе таблицы. Вы включаете настраиваемые контекстные вкладки в надстройку Office и указываете, когда они видны или скрыты, путем создания обработчиков событий, которые изменяют видимость. (Однако настраиваемые контекстные вкладки не реагируют на изменения фокуса.)

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

> [!IMPORTANT]
> Настраиваемые контекстные вкладки в настоящее время поддерживаются только в Excel и только на этих платформах и сборках.
>
> - Excel для Windows (только подписка на Microsoft 365): версия 2102 (сборка 13801.20294) или более поздняя.
> - Excel для Mac: версия 16.53.806.0 или более поздняя.
> - Excel в Интернете

> [!NOTE]
> Настраиваемые контекстные вкладки работают только на платформах, поддерживающих следующие наборы обязательных элементов. Дополнительные сведения о наборах требований и работе с ними см. в разделе "Указание приложений [Office и требований К API"](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
> - [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
>
> Вы можете использовать проверки среды выполнения в коде, чтобы проверить, поддерживает ли сочетание узла и платформы пользователя эти наборы обязательных элементов, как описано в проверках среды выполнения на наличие поддержки методов и [наборов обязательных элементов](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (Метод указания наборов обязательных элементов в манифесте, который также описан в этой статье, в настоящее время не работает для RibbonApi 1.2.) Кроме того, можно реализовать альтернативный интерфейс пользовательского интерфейса, [если настраиваемые контекстные вкладки не поддерживаются](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Поведение настраиваемых контекстных вкладок

Пользовательский интерфейс настраиваемых контекстных вкладок соответствует шаблону встроенных контекстных вкладок Office. Ниже приведены основные принципы размещения настраиваемых контекстных вкладок.

- Когда пользовательская контекстная вкладка отображается, она отображается в правой части ленты.
- Если одна или несколько встроенных контекстных вкладок и одна или несколько настраиваемых контекстных вкладок из надстроек отображаются одновременно, настраиваемые контекстные вкладки всегда находятся справа от всех встроенных контекстных вкладок.
- Если надстройка содержит несколько контекстных вкладок и есть контексты, в которых отображается несколько элементов, они отображаются в том порядке, в котором они определены в надстройке. (Направление в том же направлении, что и язык Office; то есть слева направо на языках слева направо, но справа налево на языках справа налево.) [Дополнительные сведения о том](#define-the-groups-and-controls-that-appear-on-the-tab) , как их определить, см. в разделе "Определение групп и элементов управления, которые отображаются на вкладке".
- Если несколько надстроек имеет контекстную вкладку, которая видна в определенном контексте, они отображаются в том порядке, в котором были запущены надстройки.
- *Настраиваемые контекстные* вкладки, в отличие от настраиваемых ядер вкладок, не добавляются на ленту приложения Office без возможности восстановления. Они присутствуют только в документах Office, в которых работает надстройка.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Основные шаги по добавлению контекстной вкладки в надстройку

Ниже приведены основные шаги по добавлению настраиваемой контекстной вкладки в надстройку.

1. Настройте надстройку для использования общей среды выполнения.
1. Определите вкладку, группы и элементы управления, которые отображаются на ней.
1. Зарегистрируйте контекстную вкладку в Office.
1. Укажите обстоятельства, когда вкладка будет видна.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общей среды выполнения

Чтобы добавить настраиваемые контекстные вкладки, надстройка будет использовать общую среду выполнения. Дополнительные сведения см. в разделе ["Настройка надстройки для использования общей среды выполнения"](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Определение групп и элементов управления, отображаемых на вкладке

В отличие от настраиваемых ядер вкладок, которые определены с помощью XML в манифесте, настраиваемые контекстные вкладки определяются во время выполнения с помощью большого двоичного объекта JSON. Код анализирует большой двоичный объект в объект JavaScript, а затем передает объект [методу Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . Настраиваемые контекстные вкладки присутствуют только в документах, на которых в настоящее время работает надстройка. Это отличается от настраиваемых ядер вкладок, которые добавляются на ленту приложения Office при установке надстройки и остаются доступными при открытии другого документа. Кроме того, `requestCreateControls` метод может выполняться только один раз в сеансе надстройки. При повторном вызове возникает ошибка.

> [!NOTE]
> Структура свойств и вложенных свойств большого двоичного объекта JSON (и имен ключей) примерно параллельна структуре элемента [CustomTab](/javascript/api/manifest/customtab) и его потомков в XML-коде манифеста.

Мы создадим пример контекстных вкладок большого двоичного объекта JSON пошаговые инструкции. Полная схема контекстной вкладки JSON находится на [сайте dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Если вы работаете в Visual Studio Code, этот файл можно использовать для получения IntelliSense и проверки JSON. Дополнительные сведения см. в разделе ["Изменение JSON с Visual Studio Code - схемы и параметры JSON"](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Начните с создания строки JSON с двумя свойствами массива с именем и `actions` `tabs`. Массив `actions` представляет собой спецификацию всех функций, которые могут выполняться элементами управления на контекстной вкладке. Массив `tabs` определяет одну или несколько контекстных вкладок не более *20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Этот простой пример контекстной вкладки будет содержать только одну кнопку и, следовательно, только одно действие. Добавьте следующий код в качестве единственный член массива `actions` . Обратите внимание на эту разметку:

    - Свойства `id` и `type` свойства являются обязательными.
    - Значением может `type` быть ExecuteFunction или ShowTaskpane.
    - Свойство `functionName` используется только в том случае, если значение равно `type` .`ExecuteFunction` Это имя функции, определенной в FunctionFile. Дополнительные сведения о FunctionFile см. в разделе ["Основные понятия для команд надстроек"](add-in-commands.md).
    - На следующем шаге вы сопоставляете это действие с кнопкой на контекстной вкладке.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. Добавьте следующий код в качестве единственный член массива `tabs` . Обратите внимание на эту разметку:

    - Свойство `id` является обязательным. Используйте краткий описательный идентификатор, уникальный среди всех контекстных вкладок в надстройке.
    - Свойство `label` является обязательным. Это у пользователей строка, которая служит меткой контекстной вкладки.
    - Свойство `groups` является обязательным. Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен содержать по крайней мере один член *и не более 20*. (Кроме того, существуют ограничения на количество элементов управления, которые можно использовать на настраиваемой контекстной вкладке, а также количество групп. Дополнительные сведения см. на следующем шаге.)

    > [!NOTE]
    > Объект табуляции также может `visible` иметь необязательное свойство, указывающее, отображается ли вкладка сразу же при запуске надстройки. Так как контекстные вкладки обычно скрыты до тех пор, пока событие пользователя не активирует их видимость (например, пользователь выбирает сущность определенного типа в документе), `visible` `false` свойство по умолчанию имеет значение, если оно отсутствует. В следующем разделе мы покажем, как `true` задать свойство в ответ на событие.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. В простом непрерывном примере контекстная вкладка имеет только одну группу. Добавьте следующий код в качестве единственный член массива `groups` . Обратите внимание на эту разметку:

    - Все свойства являются обязательными.
    - Свойство `id` должно быть уникальным среди всех групп в манифесте. Используйте краткий описательный идентификатор до 125 символов.
    - Является `label` понятной строкой для использования в качестве метки группы.
    - Значение `icon` свойства представляет собой массив объектов, указывающих значки, которые группа будет иметь на ленте в зависимости от размера ленты и окна приложения Office.
    - Значение `controls` свойства представляет собой массив объектов, указывающих кнопки и меню в группе. Должен быть по крайней мере один.

    > [!IMPORTANT]
    > *Общее число элементов управления на всей вкладке не может превышать 20.* Например, у вас может быть 3 группы с 6 элементами управления и четвертая группа с 2 элементами управления, но у вас не может быть 4 группы с 6 элементами управления.  

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

1. Каждая группа должна иметь значок по крайней мере двух размеров: 32x32 px и 80x80 px. Кроме того, можно использовать значки размеров 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px и 64x64 px. Office определяет, какой значок использовать в зависимости от размера ленты и окна приложения Office. Добавьте следующие объекты в массив значков. (Если размер окна и ленты достаточно велик для отображения хотя бы одного из элементов  управления в группе, значок группы не отображается. Например, просмотрите группу **"Стили** " на ленте Word при сжатии и развертывании окна Word.) Обратите внимание на эту разметку:

    - Оба свойства являются обязательными.
    - Единица `size` измерения свойства — пиксели. Значки всегда квадратные, поэтому число равно высоте и ширине.
    - Свойство `sourceLocation` указывает полный URL-адрес значка.

    > [!IMPORTANT]
    > Как и при переходе с разработки на рабочую (например, изменение домена с localhost на contoso.com) URL-адреса в манифесте надстройки, необходимо также изменить URL-адреса на контекстных вкладок JSON.

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

1. В нашем простом непрерывном примере группа имеет только одну кнопку. Добавьте следующий объект в качестве единственный член массива `controls` . Обратите внимание на эту разметку:

    - Все свойства, за исключением `enabled`, являются обязательными.
    - `type` указывает тип элемента управления. Значения могут быть "Button", "Menu" или "MobileButton".
    - `id` может содержать до 125 символов.
    - `actionId` должен быть идентификатором действия, определенного в массиве `actions` . (См. шаг 1 этого раздела.)
    - `label` является понятной строкой для использования в качестве метки кнопки.
    - `superTip` представляет расширенную форму подсказки. И свойства `title` `description` , и свойства являются обязательными.
    - `icon` указывает значки для кнопки. Здесь также применимы предыдущие примечания о значке группы.
    - `enabled` (необязательно) указывает, включена ли кнопка при запуске контекстной вкладки. Значение по умолчанию, если его нет `true`.

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

Контекстная вкладка регистрируется в Office путем вызова метода [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . Обычно это делается либо в назначенной `Office.initialize` функции, либо с помощью функции `Office.onReady` . Дополнительные сведения об этих функциях и инициализации надстройки см. в разделе ["Инициализация надстройки Office"](../develop/initialize-add-in.md). Однако метод можно вызвать в любое время после инициализации.

> [!IMPORTANT]
> Метод `requestCreateControls` может вызываться только один раз в заданный сеанс надстройки. При повторном вызове возникает ошибка.

Ниже приведен пример. Обратите внимание, что перед передачей в функцию JavaScript строку JSON необходимо преобразовать в объект JavaScript `JSON.parse` с помощью метода.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Укажите контексты, когда вкладка будет отображаться с помощью requestUpdate

Как правило, настраиваемая контекстная вкладка должна отображаться, когда инициированное пользователем событие изменяет контекст надстройки. Рассмотрим сценарий, в котором вкладка должна быть видна только при активации диаграммы (на листе excel по умолчанию).

Начните с назначения обработчиков. Обычно это делается `Office.onReady` в функции, как показано в следующем примере, которая назначает обработчики (созданные на следующем шаге) `onActivated` `onDeactivated` всем диаграммам на листе и событиям.

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

Затем определите обработчики. Ниже приведен простой `showDataTab`пример ошибки [HostRestartNeeded](#handle-the-hostrestartneeded-error) , но более надежную версию функции см. далее в этой статье. Вот что нужно знать об этом коде:

- Office определяет время обновления состояния ленты. Метод  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) ставит запрос на обновление в очередь. Метод разрешит объект `Promise` сразу после того, как запрос будет поставлен в очередь, а не при фактическом обновлении ленты.
- `requestUpdate` Параметром метода является объект [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata), который (1) задает вкладку по идентификатору точно так, как указано в *JSON*, и (2) указывает видимость вкладки.
- Если у вас несколько настраиваемых контекстных вкладок, которые должны отображаться в одном контексте, просто добавьте в массив дополнительные объекты табуляции `tabs` .

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

Обработчик для скрытия вкладки практически идентичен, за исключением того, `visible` что он задает для свойства обратно значение `false`.

Библиотека JavaScript для Office также предоставляет несколько интерфейсов (типов), упрощая создание`RibbonUpdateData` объекта. Ниже приведена функция `showDataTab` в TypeScript, которая использует эти типы.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Одновременное переключение видимости вкладки и состояния включенной кнопки

Этот `requestUpdate` метод также используется для переключения состояния включенной или отключенной пользовательской кнопки на настраиваемой контекстной вкладке или настраиваемой вкладке ядра. Дополнительные сведения об этом см. в разделе "Включение и отключение [команд надстроек"](disable-add-in-commands.md). Возможны сценарии, в которых требуется одновременно изменить видимость вкладки и состояние включенной кнопки. Это можно сделать одним вызовом .`requestUpdate` Ниже приведен пример, в котором кнопка на основной вкладке включена одновременно с видимой контекстной вкладкой.

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

В следующем примере включенная кнопка находится на той же контекстной вкладке, которая отображается.

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

## <a name="open-a-task-pane-from-contextual-tabs"></a>Открытие области задач с контекстных вкладок

Чтобы открыть область задач с кнопки на настраиваемой контекстной вкладке, создайте действие в JSON с и `type` `ShowTaskpane`. Затем определите кнопку со свойством `actionId` , заданным для `id` действия. Откроется область задач по умолчанию, указанная элементом **\<Runtime\>** манифеста.

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

Чтобы открыть любую область задач, которая не является областью задач по умолчанию, `sourceLocation` укажите свойство в определении действия. В следующем примере вторая область задач открывается с другой кнопки.

> [!IMPORTANT]
>
> - Если для `sourceLocation` действия указано значение a, область задач не *использует* общую среду выполнения. Он выполняется в новой среде выполнения JavaScript.
> - Не более одной области задач могут использовать общую среду выполнения, `ShowTaskpane` поэтому не более одного действия типа могут опустить `sourceLocation` это свойство.

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

Переданный большой `requestCreateControls` двоичный объект JSON не локализован так же, как локализованная разметка манифеста для настраиваемых ядер вкладок (что описано в разделе "Локализация элемента управления" из [манифеста](../develop/localization.md#control-localization-from-the-manifest)). Вместо этого локализация должна выполняться во время выполнения с использованием отдельных больших двоичных объектов JSON для каждого языкового стандарта. Рекомендуется использовать инструкцию, `switch` которая проверяет свойство [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) . Ниже приведен пример.

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

Затем код вызывает функцию для получения `requestCreateControls`локализованного большого двоичного объекта, который передается, как показано в следующем примере.

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Рекомендации для настраиваемых контекстных вкладок

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Реализация альтернативного интерфейса пользовательского интерфейса, если настраиваемые контекстные вкладки не поддерживаются

Некоторые сочетания платформы, приложения Office и сборки Office не поддерживаются `requestCreateControls`. Ваша надстройка должна быть разработана для предоставления альтернативного интерфейса пользователям, которые запускают надстройку в одном из этих сочетаний. В следующих разделах описаны два способа предоставления резервного интерфейса.

#### <a name="use-noncontextual-tabs-or-controls"></a>Использование неконтекстуальных вкладок или элементов управления

Существует элемент манифеста [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi), который предназначен для создания резервного интерфейса в надстройке, которая реализует настраиваемые контекстные вкладки, когда надстройка выполняется на приложении или платформе, не поддерживающих настраиваемые контекстные вкладки.

Простейшая стратегия использования этого элемента заключается в *том, чтобы* определить настраиваемую ядровую вкладку (то есть неконтекстовую настраиваемую вкладку) в манифесте, которая дублирует настройки ленты настраиваемых контекстных вкладок в надстройке. Но вы добавляете `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` в качестве первого дочернего элемента повторяющихся элементов [group](/javascript/api/manifest/group), [Control](/javascript/api/manifest/control) и menu **\<Item\>** на настраиваемых ядрах вкладок. Это может быть следующим образом.

- Если надстройка работает в приложении и платформе, поддерживающих настраиваемые контекстные вкладки, пользовательские основные группы и элементы управления не будут отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет создана, когда надстройка вызывает `requestCreateControls` метод.
- Если надстройка работает в приложении  `requestCreateControls`или платформе, которые не поддерживаются, элементы отображаются на вкладке "Настраиваемое ядро".

Ниже приведен пример. Обратите внимание, что MyButton будет отображаться на настраиваемой вкладке ядра только в том случае, если настраиваемые контекстные вкладки не поддерживаются. Но родительская группа и настраиваемая вкладка ядра будут отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки.

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

Дополнительные примеры см. в [разделе OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi).

Если родительская `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`группа или меню помечены, она не отображается, а вся ее дочерняя разметка игнорируется, если настраиваемые контекстные вкладки не поддерживаются. Поэтому не имеет значения, имеет ли какой-либо из этих **\<OverriddenByRibbonApi\>** дочерних элементов элемент или его значение. Это означает, что если элемент меню или элемент управления должен быть виден во всех контекстах, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`он не только не должен быть помечен, но и его предок и группа также не должны быть помечены таким *образом*.

> [!IMPORTANT]
> Не *помечать все* дочерние элементы группы или меню .`<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` Это не имеет смысла, если родительский элемент `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` помечен по причинам, указанным в предыдущем абзаце. Кроме того, **\<OverriddenByRibbonApi\>** если оставить родительский элемент ( `false`или задать для него значение), родительский элемент будет отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки, но он будет пустым, если они поддерживаются. Таким образом, если все дочерние элементы не должны отображаться при поддерживаемых настраиваемых контекстных вкладок, пометьте родительский элемент `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Использование интерфейсов API, которые отображают или скрывают область задач в указанных контекстах

В качестве альтернативы надстройка **\<OverriddenByRibbonApi\>** может определить область задач с элементами управления пользовательского интерфейса, которые дублируют функции элементов управления на настраиваемой контекстной вкладке. Затем используйте методы [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) и [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) , чтобы отобразить область задач, когда была бы показана контекстная вкладка, если она поддерживалась. Дополнительные сведения об использовании этих методов см. в разделе "Показать или скрыть область задач" [надстройки Office](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Обработка ошибки HostRestartNeeded

В некоторых случаях Office не может обновить ленту и возвращает ошибку. Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть. Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`. Код должен обработать эту ошибку. Ниже приведен пример того, как это сделать. В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.

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

- [Пример кода. Создание настраиваемых контекстных вкладок на ленте](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- Демонстрация примера контекстных вкладок в сообществе

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]
