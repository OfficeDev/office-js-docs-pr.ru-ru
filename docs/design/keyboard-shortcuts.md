---
title: Настраиваемые сочетания клавиш в надстройки Office
description: Узнайте, как добавить в надстройку Office пользовательские сочетания клавиш, также известные как сочетания клавиш.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: dc99674b92ebb415b1d49fb28821d8c2e34c8077
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789151"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Добавление пользовательских сочетания клавиш в надстройки Office (предварительная версия)

Сочетания клавиш, также известные как сочетания клавиш, позволяют пользователям надстройки работать эффективнее и улучшают ее доступность для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Чтобы начать с рабочей версии надстройки с уже включенными сочетаниями клавиш, клонировать и запустить пример сочетания клавиш [Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) Когда вы будете готовы добавить сочетания клавиш в вашу собственную надстройку, продолжайте работу с этой статьей.

Чтобы добавить сочетания клавиш в надстройку, необходимо три шага:

1. [Настройте манифест надстройки.](#configure-the-manifest)
1. [Создайте или отредактируетЕ JSON-файл](#create-or-edit-the-shortcuts-json-file) ярлыков, чтобы определить действия и их сочетания клавиш.
1. [Добавьте один или несколько вызовов](#create-a-mapping-of-actions-to-their-functions) API [Office.actions.associate](/javascript/api/office/office.actions#associate) во время работы, чтобы соотоставить функцию с каждым действием.

## <a name="configure-the-manifest"></a>Настройка манифеста

Манифесту необходимо внести два небольших изменения. Первый — позволить надстройки использовать общую времени работы, а другой — указать файл в формате JSON, в котором вы определили сочетания клавиш.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общей времени работы

Для добавления настраиваемого сочетания клавиш надстройка использует общую времени работы. Для получения дополнительных [сведений настройте надстройку](../develop/configure-your-add-in-to-use-a-shared-runtime.md)для использования общей времени работы.

### <a name="link-the-mapping-file-to-the-manifest"></a>Привязка файла сопоставления к манифесту

Непосредственно *под* элементом манифеста (не внутри) добавьте `<VersionOverrides>` элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Установите для атрибута полный URL-адрес JSON-файла в проекте, который будет создан на `Url` более позднем этапе.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Создание или изменение JSON-файла ярлыков

Создайте JSON-файл в проекте. Убедитесь, что путь к файлу соответствует расположению, указанному для `Url` атрибута [элемента ExtendedOverrides.](../reference/manifest/extendedoverrides.md) В этом файле описываются сочетания клавиш и действия, которые они будут вызывать.

1. Внутри JSON-файла есть два массива. Массив действий будет содержать объекты, которые определяют действия, которые необходимо вызвать, а массив ярлыков будет содержать объекты, которые соотнося сочетания клавиш с действиями. Вот пример:

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

    Дополнительные сведения об объектах JSON см. в подстройки объектов [действий](#constructing-the-action-objects) и создания объектов [ярлыков.](#constructing-the-shortcut-objects) Полная схема для ярлыков JSON находится в [extended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

    > [!NOTE]
    > В этой статье вы можете использовать control, а не CTRL.

    На более позднем этапе эти действия будут сами сописываться с функциями, которые вы пишете. В этом примере вы позже созовем SHOWTASKPANE с функцией, которая вызывает метод, и HIDETASKPANE с функцией, которая `Office.addin.showAsTaskpane` вызывает `Office.addin.hide` метод.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Создание сопоставления действий с их функциями

1. В проекте откройте файл JavaScript, загруженный HTML-страницей в `<FunctionFile>` элементе.
1. В файле JavaScript используйте API [Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотоставить каждое действие, указанное в файле JSON, с функцией JavaScript. Добавьте в файл следующий javaScript. Обратите внимание на следующие вопросы о коде:

    - Первый параметр — это одно из действий из JSON-файла.
    - Второй параметр — это функция, которая запускается, когда пользователь нажмет сочетание клавиш, которое со карты с действием в JSON-файле.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Чтобы продолжить пример, используйте `'SHOWTASKPANE'` его в качестве первого параметра.
1. Для тела функции используйте метод [Office.addin.showTaskpane,](/javascript/api/office/office.addin#showastaskpane--) чтобы открыть область задач надстройки. После этого код должен выглядеть следующим образом:

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

1. Добавьте второй вызов `Office.actions.associate` функции, чтобы соединить действие `HIDETASKPANE` с функцией, которая вызывает [Office.addin.hide.](/javascript/api/office/office.addin#hide--) Ниже приведен пример.

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

После предыдущих действий надстройка может переключать видимость области задач, нажимая **клавиши CTRL+SHIFT+СТРЕЛКА** ВВЕРХ и **CTRL+SHIFT+СТРЕЛКА ВНИЗ.** Это поведение такое же, как показано в примере надстройки [Excel с сочетаниями клавиш.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)

## <a name="details-and-restrictions"></a>Сведения и ограничения

### <a name="constructing-the-action-objects"></a>Создание объектов действий

При указании объектов в массиве shortcuts.jsиспользуйте `action` следующие рекомендации:

- Имена свойств `id` являются `name` обязательными.
- Свойство используется для уникальной идентификации `id` действия, вызываемого с помощью сочетания клавиш.
- Свойство `name` должно быть пользовательской строкой, описывающий действие. Это должно быть сочетание символов A - Z, a - z, 0 - 9 и знаков препинания "-", "_" и "+".
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

Полная схема для ярлыков JSON находится в [extended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

### <a name="constructing-the-shortcut-objects"></a>Создание объектов ярлыков

При указании объектов в массиве shortcuts.jsиспользуйте `shortcuts` следующие рекомендации:

- Имена свойств `action` , и являются `key` `default` обязательной.
- Значение свойства является строкой и должно соответствовать одному из `action` `id` свойств в объекте действия.
- Свойство может быть любым сочетанием символов `default` A - Z, a -z, 0 -9 и знаками препинания "-", "_" и "+". (По соглашению буквы нижнего дела не используются в этих свойствах.)
- Свойство должно содержать имя по крайней мере одного ключа модификатора `default` (ALT, CTRL, SHIFT) и только одного другого ключа.
- Для Mac также поддерживается клавиша модификатора COMMAND.
- Для Компьютеров Mac ALT со картой OPTION. Для Windows command со карты с клавишей CTRL.
- Если два символа связаны с одним и тем же физическим ключом в стандартной клавиатуре, они являются синонимами в свойстве; например, ALT+a и ALT+A — это один и тот же ярлык, как `default` и CTRL+- и CTRL+, так как "-" и "_" являются одним и тем же физическим \_ ключом.
- Символ "+" указывает, что клавиши с обеих сторон нажимаются одновременно.

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

Полная схема для ярлыков JSON находится в [extended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

> [!NOTE]
> Клавиши, также известные как последовательное сочетания клавиш, например ярлык Excel для выбора цвета заливки **ALT+H, H,** не поддерживаются в надстройки Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Использование ярлыков, когда фокус находится в области задач

В настоящее время сочетания клавиш для надстройки Office можно вызывать только в том случае, если фокус пользователя находится на этом планшете. Если фокус пользователя находится в пользовательском интерфейсе Office (например, в области задач), ни один из ярлыков надстройки не игнорируется. В качестве обходного решения надстройка может определять обработчики клавиатуры, которые могут вызывать определенные действия, когда фокус пользователя находится в пользовательском интерфейсе надстройки.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Использование сочетаний клавиш, которые уже используются Office или другой надстройки

Во время предварительного просмотра не существует системы для определения того, что происходит, когда пользователь нажимает сочетание клавиш, которое зарегистрировано надстройки, а также Office или другой надстройки. Поведение неопределяется.

В настоящее время обходной путь не существует, если две или более надстроек зарегистрировали один и тот же ярлык клавиатуры, но вы можете свести к минимуму конфликты с Excel, выполив указанные здесь действия.

- Используйте в надстройки только сочетания клавиш со следующим шаблоном: **CTRL+SHIFT+ALT+* x***, где *x* — это другой ключ.
- Если вам нужны дополнительные сочетания клавиш, проверьте список сочетания клавиш [Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)и избегайте их использования в надстройки.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Ярлыки браузера, которые не могут быть переопределены

Нельзя использовать любое из следующих сочетаний клавиатуры. Они используются браузерами и не могут быть переопределены. Этот список является ходом выполнения. Если вы обнаружите другие сочетания, которые не могут быть переопределены, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.

- CTRL+N
- CTRL+SHIFT+N
- CTRL+T
- CTRL+SHIFT+T
- CTRL+W
- CTRL+PgUp/PgDn

## <a name="next-steps"></a>Дальнейшие действия

- См. пример надстройки [Excel-keyboard-shortcuts.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)
