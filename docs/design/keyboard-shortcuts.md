---
title: Пользовательские ярлыки клавиатуры в надстройки Office
description: Узнайте, как добавить в надстройку Office настраиваемые клавиши, также известные как сочетания клавиш.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: c767c6d5bc23f0a44422452839cd8bdf87bd8715
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505201"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Добавление ярлыков настраиваемой клавиатуры в надстройки Office (предварительный просмотр)

Ярлыки клавиатуры, также известные как сочетания клавиш, позволяют пользователям надстройки работать эффективнее и они улучшают доступность надстройки для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Чтобы начать с рабочей версии надстройки с уже включенными клавишами, клонировать и запускать примеры ярлыков [клавиатуры Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) Если вы готовы добавить ярлыки клавиатуры в собственную надстройку, продолжи эту статью.

Существует три шага, чтобы добавить в надстройку ярлыки клавиатуры:

1. [Настройка манифеста надстройки.](#configure-the-manifest)
1. [Создание или изменение ярлыков JSON-файла для](#create-or-edit-the-shortcuts-json-file) определения действий и их клавиш.
1. [Добавьте один или несколько вызовов](#create-a-mapping-of-actions-to-their-functions) [API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотоставить функцию с каждым действием.

## <a name="configure-the-manifest"></a>Настройка манифеста

В манифест необходимо внести два небольших изменения. Один из них — включить надстройку для использования общего времени работы, а другой — указать на файл в формате JSON, в котором определены ярлыки клавиатуры.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общего времени работы

Добавление пользовательских ярлыков клавиатуры требует от надстройки использовать общее время работы. Дополнительные сведения: [Настройка надстройки для использования общего времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

### <a name="link-the-mapping-file-to-the-manifest"></a>Привязка файла сопоставления к манифесту

Сразу *ниже* (не внутри) элемента `<VersionOverrides>` манифеста добавьте элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Установите атрибут для полного URL-адреса файла JSON в проекте, который будет создан `Url` на более позднем этапе.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Создание или изменение ярлыков JSON-файла

Создайте файл JSON в проекте. Убедитесь, что путь файла соответствует расположению, указанному для атрибута элемента `Url` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) В этом файле будут описаны ярлыки клавиатуры и действия, которые они будут вызывать.

1. В файле JSON есть два массива. Массив действий будет содержать объекты, которые определяют действия, которые будут вызываться, а массив ярлыков будет содержать объекты, которые соотносят комбинации ключей с действиями. Вот пример:

    ```json
    {
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    }
    ```

    Дополнительные сведения об объектах JSON см. в дополнительных сведениях [о](#constructing-the-action-objects) том, как создавать объекты действия и создавать [объекты ярлыка.](#constructing-the-shortcut-objects) Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

    > [!NOTE]
    > В этой статье можно использовать "CONTROL" на месте "CTRL".

    На более позднем этапе действия сами будут соедему с функциями, которые вы пишете. В этом примере вы позже назовет SHOWTASKPANE функцией, которая вызывает метод, а HIDETASKPANE — функцией, которая `Office.addin.showAsTaskpane` вызывает `Office.addin.hide` метод.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Создание сопоставления действий с их функциями

1. В проекте откройте файл JavaScript, загруженный вашей htmL-страницей в `<FunctionFile>` элементе.
1. В файле JavaScript используйте [API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы составить карту каждого действия, указанного в файле JSON, с функцией JavaScript. Добавьте в файл следующий JavaScript. Обратите внимание на следующее:

    - Первый параметр — это одно из действий из файла JSON.
    - Второй параметр — это функция, которая выполняется при нажатии клавиши на сочетание ключей, относясь к действию в файле JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Чтобы продолжить пример, используйте `'SHOWTASKPANE'` в качестве первого параметра.
1. Чтобы открыть область задач надстройки, используйте метод [Office.addin.showTaskpane.](/javascript/api/office/office.addin#showastaskpane--) После этого код должен выглядеть следующим образом:

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. Добавьте второй вызов `Office.actions.associate` функции, чтобы соединить действие с функцией, вызываемой `HIDETASKPANE` [Office.addin.hide.](/javascript/api/office/office.addin#hide--) Ниже приведен пример.

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

После предыдущих действий надстройка позволяет переключать видимость области задач, нажав клавишу **стрелки Ctrl+Shift+Up** и клавишу **стрелки Ctrl+Shift+Down.** Это такое же поведение, как показано в примере надстройки [клавиш Excel клавиатуры](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).

## <a name="details-and-restrictions"></a>Сведения и ограничения

### <a name="constructing-the-action-objects"></a>Построение объектов действия

Используйте следующие рекомендации при указании объектов в массиве `action` shortcuts.js:

- Имена свойств `id` и `name` обязательны.
- Свойство `id` используется для уникальной идентификации действия, вызываемого с помощью ярлыка клавиатуры.
- Свойство `name` должно быть удобной строкой, описываемой действием. Это должно быть сочетание символов A - Z, a - z, 0 - 9 и знаков препинания "-", "_" и "+".
- Свойство `type`— необязательное. В `ExecuteFunction` настоящее время поддерживается только тип.

Ниже приведен пример.

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

### <a name="constructing-the-shortcut-objects"></a>Построение объектов ярлыка

Используйте следующие рекомендации при указании объектов в массиве `shortcuts` shortcuts.js:

- Имена свойств `action` `key` и `default` обязательно.
- Значение свойства является строкой и должно соответствовать одному из свойств `action` `id` объекта действия.
- Свойство может быть любым сочетанием символов A - Z, a -z, 0 - 9, а знаки препинания `default` "-", "_" и "+". (По соглашению в этих свойствах не используются буквы более низкого уровня.)
- Свойство должно содержать имя по крайней мере одного ключа модификатора `default` (ALT, CTRL, SHIFT) и только одного ключа.
- Для macs мы также поддерживаем ключ модификатора COMMAND.
- Для Компьютеров Mac ALT соедем на клавишу OPTION. Для Windows командная команда относит к клавише CTRL.
- Если два символа связаны с одним и тем же физическим ключом в стандартной клавиатуре, они являются синонимами в свойстве; например, ALT+a и ALT+A являются одним и тем же ярлыком, как `default` и CTRL+- и CTRL+, так как "-" и "_" являются одним и тем же физическим \_ ключом.
- Символ "+" указывает, что клавиши с обеих сторон нажаты одновременно.

Ниже приведен пример.

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+UP"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "CTRL+SHIFT+DOWN"
            }
        }
    ]
```

Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

> [!NOTE]
> Клавиши, также известные как последовательное клавиши, например ярлык Excel для выбора цвета заполнения **Alt+H, H,** не поддерживаются в надстройки Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Использование ярлыков, когда фокус находится в области задач

В настоящее время ярлыки клавиатуры для надстройки Office можно вызывать только в том случае, если фокус пользователя находится в таблице. Если фокус пользователя находится внутри пользовательского интерфейса Office (например, области задач), ни один из ярлыков надстройки не игнорируется. В качестве обхода надстройка может определять обработчики клавиатуры, которые могут вызывать определенные действия, когда фокус пользователя находится внутри пользовательского интерфейса надстройки.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Использование комбинаций ключей, которые уже используются Office или другой надстройки

Во время предварительного просмотра не существует системы определения того, что происходит, когда пользователь нажимает клавишу, зарегистрированную надстройка, а также Office или другой надстройки. Поведение неопределяется.

В настоящее время не существует обхода, когда два или несколько надстроек зарегистрировали один и тот же ярлык клавиатуры, но можно минимизировать конфликты с Excel с помощью этих методов:

- Используйте только клавиши со следующим шаблоном в надстройки: **Ctrl+Shift+Alt+* x***, где *x* — это другой ключ.
- Если вам нужно больше клавиш, проверьте список ярлыков [клавиатуры Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)и не использовать их в надстройки.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Ярлыки браузера, которые нельзя переопределять

Вы не можете использовать ни одной из следующих комбинаций клавиатуры. Они используются браузерами и не могут быть переопределены. Этот список находится в процессе выполнения. Если вы обнаружите другие комбинации, которые нельзя переопределять, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="localize-the-keyboard-shortcuts-json"></a>Локализовать ярлыки клавиатуры JSON

Если надстройка поддерживает несколько локалов, необходимо локализовать свойство `name` объектов действия. Кроме того, если в любом из локаутах, поддерживаюх надстройку, есть алфавиты или различные системы записи, а значит, и другие клавиатуры, возможно, потребуется также локализовать ярлыки. Сведения о том, как локализовать клавиши ярлыков JSON, см. в рубрезе [Localize extended overrides.](../develop/localization.md#localize-extended-overrides)

## <a name="next-steps"></a>Дальнейшие действия

- См. пример [надстройки excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).
- Получите обзор работы с расширенными переопределениями в Работе с расширенными [переопределениями манифеста.](../develop/extended-overrides.md)
