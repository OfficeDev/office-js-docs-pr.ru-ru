---
title: Настраиваемые сочетания клавиш в надстройках Office
description: Узнайте, как добавить в надстройку Office пользовательские сочетания клавиш, которые также называются сочетаниями клавиш.
ms.date: 11/09/2020
localization_priority: Normal
ms.openlocfilehash: 40009dd92787b7c220bb8cfc741cffb2e4b68a9e
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132041"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Добавление настраиваемых сочетаний клавиш в надстройки Office (Предварительная версия)

Сочетания клавиш, называемые также сочетаниями клавиш, позволяют пользователям вашей надстройки работать эффективнее и расширять возможности надстройки для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Чтобы начать работу с рабочей версией надстройки с включенными сочетаниями клавиш, выполните клонирование и выполните примеры сочетаний [клавиш Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts). Когда вы будете готовы добавить сочетания клавиш для своей надстройки, перейдите к этой статье.

Добавление сочетаний клавиш в надстройку состоит из трех этапов:

1. [Настройте манифест надстройки](#configure-the-manifest).
1. [Создайте или измените файл ярлыков JSON](#create-or-edit-the-shortcuts-json-file) , чтобы определить действия и их сочетания клавиш.
1. [Добавьте один или несколько вызовов среды выполнения](#create-a-mapping-of-actions-to-their-functions) API [Office. Actions.](/javascript/api/office/office.actions#associate) Map, чтобы сопоставить функцию с каждым действием.

## <a name="configure-the-manifest"></a>Настройка манифеста

В манифесте есть два небольших изменения, которые необходимо выполнить. Один — позволить надстройке использовать общую среду выполнения, а другая — указать на файл в формате JSON, в котором были определены сочетания клавиш.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общей среды выполнения

Добавление пользовательских сочетаний клавиш требует, чтобы ваша надстройка использовала общую среду выполнения. Для получения дополнительных сведений [Настройте надстройку для использования общей среды выполнения](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Связывание файла сопоставления с манифестом

Непосредственно *ниже* (не внутри) `<VersionOverrides>` элемента в манифесте добавьте элемент [екстендедоверридес](../reference/manifest/extendedoverrides.md) . Присвойте `Url` атрибуту полный URL-адрес JSON-файла в проекте, который будет создан на более позднем этапе.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Создание или редактирование файла ярлыков JSON

Создайте файл JSON в проекте. Убедитесь, что путь к файлу совпадает с расположением, указанным для `Url` атрибута элемента [екстендедоверридес](../reference/manifest/extendedoverrides.md) . В этом файле будут описаны сочетания клавиш и действия, которые они будут вызывать.

1. В файле JSON существует два массива. Массив Actions будет содержать объекты, определяющие действия, которые необходимо вызвать, а массив ярлыков будет содержать объекты, которые сопоставлены с сочетаниями клавиш на действия. Вот пример:

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

    Дополнительные сведения об объектах JSON приведены в статье [Создание объектов Action](#constructing-the-action-objects) и [Создание объектов ярлыков](#constructing-the-shortcut-objects). Полная схема для ярлыков JSON [extended-manifest.schema.jsвключена](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json). (Примечание: ссылка на схему может не работать рано в периоде предварительной версии.)

    > [!NOTE]
    > В этой статье можно использовать элемент управления вместо "CTRL".

    На более позднем этапе действия будут сопоставлены с написанными функциями. В этом примере позднее показано сопоставление SHOWTASKPANE с функцией, которая вызывает `Office.addin.showAsTaskpane` метод, и хидетаскпане в функцию, которая вызывает `Office.addin.hide` метод.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Создание сопоставления действий с их функциями

1. В проекте откройте файл JavaScript, загруженный HTML-страницей в `<FunctionFile>` элементе.
1. В файле JavaScript используйте API [Office. Actions.](/javascript/api/office/office.actions#associate) Map, чтобы сопоставить каждое действие, указанное в JSON-файле, с функцией JavaScript. Добавьте в файл приведенный ниже код JavaScript. Обратите внимание на следующие особенности кода:

    - Первый параметр — это одно из действий из JSON-файла.
    - Второй параметр — это функция, которая запускается, когда пользователь нажимает комбинацию клавиш, сопоставленную с действием в JSON-файле.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Чтобы продолжить пример, используйте `'SHOWTASKPANE'` в качестве первого параметра.
1. Для основной части функции используйте метод [Office. AddIn. showTaskpane](/javascript/api/office/office.addin#showastaskpane--) , чтобы открыть область задач надстройки. После завершения код должен выглядеть следующим образом:

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

1. Добавьте второй вызов `Office.actions.associate` функции, чтобы сопоставить `HIDETASKPANE` действие с функцией, которая вызывает [Office. AddIn. Hide](/javascript/api/office/office.addin#hide--). Ниже приведен пример.

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

После выполнения описанных выше действий надстройка позволяет переключать видимость области задач, нажимая клавиши **Ctrl + Shift + стрелка вверх** и **Ctrl + Shift + стрелка вниз**. Это то же поведение, которое показано в [примере надстройки "сочетания клавиш Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)".

## <a name="details-and-restrictions"></a>Сведения и ограничения

### <a name="constructing-the-action-objects"></a>Создание объектов Action

При указании объектов в массиве shortcuts.jsследует придерживаться следующих рекомендаций `action` .

- Имена свойств `id` и `name` являются обязательными.
- `id`Свойство используется для уникальной идентификации действия, которое вызывается с помощью сочетания клавиш.
- `name`Свойство должно представлять собой удобную пользователю строку, описывающую действие. Он должен быть комбинацией символов A – Z, a – z, 0-9 и знаков препинания "–", "_" и "+".
- Свойство `type`— необязательное. В настоящее время `ExecuteFunction` поддерживается только тип.

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

Полная схема для ярлыков JSON [extended-manifest.schema.jsвключена](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json). (Примечание: ссылка на схему может не работать рано в периоде предварительной версии.)

### <a name="constructing-the-shortcut-objects"></a>Создание объектов ярлыков

При указании объектов в массиве shortcuts.jsследует придерживаться следующих рекомендаций `shortcuts` .

- Имена свойств `action` `key` и `default` обязательные.
- Значение `action` свойства является строкой и должно удовлетворять одному из `id` свойств в объекте Action.
- `default`Свойство может быть любым сочетанием символов a – z, a – z, 0-9 и знаков препинания "–", "_" и "+". (В соответствии с соглашением буквы нижнего регистра не используются в этих свойствах.)
- `default`Свойство должно содержать имя по крайней мере одной клавиши-модификатора (Alt, CTRL, Shift) и только один ключ.
- Для Макинтош мы также поддерживаем клавишей CTRL COMMAND.
- Для Макинтошей атрибут ALT сопоставлен с ключом OPTION. Для Windows команда сопоставляется с клавишей CTRL.
- Если два символа связаны с одним и тем же физическим ключом на стандартной клавиатуре, то они являются синонимами в `default` свойстве, например ALT + a, а Alt + a — это одно сочетание клавиш, поэтому клавиши CTRL +-и CTRL +, \_ так как "-" и "_" являются одним и тем же физическим ключом.
- Символ "+" указывает на то, что клавиши с любой стороны объекта одновременно нажаты.

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

Полная схема для ярлыков JSON [extended-manifest.schema.jsвключена](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json). (Примечание: ссылка на схему может не работать рано в периоде предварительной версии.)

> [!NOTE]
> Подсказки, которые также называются последовательной клавишей, такие как ярлык Excel для выбора цвета заливки **ALT + H, H**, не поддерживаются в надстройках Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Использование сочетаний клавиш, когда фокус находится в области задач

В настоящее время сочетания клавиш для надстройки Office могут вызываться только в том случае, если фокус пользователя находится на листе. Когда фокус пользователя находится в пользовательском интерфейсе Office (например, область задач), ни одна из ее ярлыков не игнорируется. В качестве обходного решения надстройка может определить обработчики клавиатуры, которые могут вызывать определенные действия, когда фокус пользователя находится в пользовательском интерфейсе надстройки.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Использование сочетаний клавиш, которые уже используются в Office или другой надстройке

В течение периода предварительного просмотра нет системы для определения действий, которые происходят, когда пользователь нажимает сочетание клавиш, зарегистрированное надстройкой, а также Office или другой надстройкой. Поведение не определено.

В настоящее время не существует решения, в котором две или более надстройки зарегистрировали одну комбинацию клавиш, но вы можете минимизировать конфликты с Excel, выполнив приведенные ниже рекомендации.

- Используйте только сочетания клавиш со следующим шаблоном в надстройке: **CTRL + SHIFT + ALT +* x * * *, где *x* — это другой ключ.
- Если вам нужны дополнительные сочетания клавиш, проверьте список сочетаний [клавиш Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)и старайтесь не использовать их в вашей надстройке.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Ярлыки браузеров, которые не могут быть переопределены

Вы не можете использовать следующие сочетания клавиш. Они используются браузерами и не могут быть переопределены. Этот список является рабочим процессом. Если вы обнаружите другие сочетания, которые невозможно переопределить, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.

- Ctrl + N
- Ctrl + Shift + N
- Ctrl + T
- Ctrl + Shift + T
- Ctrl + W
- Ctrl + ПГУП/Пгдн

## <a name="next-steps"></a>Дальнейшие действия

- В этой статье приведены примеры сочетаний [клавиш Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)для надстроек.
