---
title: Настраиваемые клавиши в Office надстройки
description: Узнайте, как добавить в надстройку настраиваемые клавиши, также известные как комбинации ключей, Office надстройку.
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: de8ce0d89dca6745cba96ac9a5ea946d50d41de4
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349259"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a><span data-ttu-id="68c4b-103">Добавление настраиваемого сочетания клавиш в Office надстройки</span><span class="sxs-lookup"><span data-stu-id="68c4b-103">Add custom keyboard shortcuts to your Office Add-ins</span></span>

<span data-ttu-id="68c4b-104">Ярлыки клавиатуры, также известные как сочетания клавиш, позволяют пользователям надстройки работать более эффективно.</span><span class="sxs-lookup"><span data-stu-id="68c4b-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently.</span></span> <span data-ttu-id="68c4b-105">Ярлыки клавиатуры также улучшают доступность надстройки для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.</span><span class="sxs-lookup"><span data-stu-id="68c4b-105">Keyboard shortcuts also improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="68c4b-106">Чтобы начать с рабочей версии надстройки с уже включенными клавишами, клонировать и запускать Excel [клавиши.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="68c4b-106">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="68c4b-107">Если вы готовы добавить ярлыки клавиатуры в собственную надстройку, продолжи эту статью.</span><span class="sxs-lookup"><span data-stu-id="68c4b-107">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="68c4b-108">Существует три шага, чтобы добавить в надстройку ярлыки клавиатуры:</span><span class="sxs-lookup"><span data-stu-id="68c4b-108">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="68c4b-109">[Настройка манифеста надстройки.](#configure-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="68c4b-109">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="68c4b-110">[Создание или изменение ярлыков JSON-файла для](#create-or-edit-the-shortcuts-json-file) определения действий и их клавиш.</span><span class="sxs-lookup"><span data-stu-id="68c4b-110">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="68c4b-111">[Добавьте один или несколько вызовов](#create-a-mapping-of-actions-to-their-functions) [API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотоставить функцию с каждым действием.</span><span class="sxs-lookup"><span data-stu-id="68c4b-111">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="68c4b-112">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="68c4b-112">Configure the manifest</span></span>

<span data-ttu-id="68c4b-113">В манифест необходимо внести два небольших изменения.</span><span class="sxs-lookup"><span data-stu-id="68c4b-113">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="68c4b-114">Один из них — включить надстройку для использования общего времени работы, а другой — указать на файл в формате JSON, в котором определены ярлыки клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68c4b-114">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="68c4b-115">Настройка надстройки для использования общего времени работы</span><span class="sxs-lookup"><span data-stu-id="68c4b-115">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="68c4b-116">Добавление пользовательских ярлыков клавиатуры требует от надстройки использовать общее время работы.</span><span class="sxs-lookup"><span data-stu-id="68c4b-116">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="68c4b-117">Дополнительные сведения: [Настройка надстройки для использования общего времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="68c4b-117">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="68c4b-118">Привязка файла сопоставления к манифесту</span><span class="sxs-lookup"><span data-stu-id="68c4b-118">Link the mapping file to the manifest</span></span>

<span data-ttu-id="68c4b-119">Сразу *ниже* (не внутри) элемента `<VersionOverrides>` манифеста добавьте элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="68c4b-119">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="68c4b-120">Установите атрибут для полного URL-адреса файла JSON в проекте, который будет создан `Url` на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="68c4b-120">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="68c4b-121">Создание или изменение ярлыков JSON-файла</span><span class="sxs-lookup"><span data-stu-id="68c4b-121">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="68c4b-122">Создайте файл JSON в проекте.</span><span class="sxs-lookup"><span data-stu-id="68c4b-122">Create a JSON file in your project.</span></span> <span data-ttu-id="68c4b-123">Убедитесь, что путь файла соответствует расположению, указанному для атрибута элемента `Url` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="68c4b-123">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="68c4b-124">В этом файле будут описаны ярлыки клавиатуры и действия, которые они будут вызывать.</span><span class="sxs-lookup"><span data-stu-id="68c4b-124">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="68c4b-125">В файле JSON есть два массива.</span><span class="sxs-lookup"><span data-stu-id="68c4b-125">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="68c4b-126">Массив действий будет содержать объекты, которые определяют действия, которые будут вызываться, а массив ярлыков будет содержать объекты, которые соотносят комбинации ключей с действиями.</span><span class="sxs-lookup"><span data-stu-id="68c4b-126">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="68c4b-127">Пример:</span><span class="sxs-lookup"><span data-stu-id="68c4b-127">Here is an example:</span></span>

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
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    <span data-ttu-id="68c4b-128">Дополнительные сведения об объектах JSON см. в дополнительных сведениях [о конструкторе](#construct-the-action-objects) объектов действий [и создания объектов ярлыка.](#construct-the-shortcut-objects)</span><span class="sxs-lookup"><span data-stu-id="68c4b-128">For more information about the JSON objects, see [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects).</span></span> <span data-ttu-id="68c4b-129">Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="68c4b-129">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="68c4b-130">В этой статье можно использовать "CONTROL" на месте "Ctrl".</span><span class="sxs-lookup"><span data-stu-id="68c4b-130">You can use "CONTROL" in place of "Ctrl" throughout this article.</span></span>

    <span data-ttu-id="68c4b-131">На более позднем этапе действия сами будут соедему с функциями, которые вы пишете.</span><span class="sxs-lookup"><span data-stu-id="68c4b-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="68c4b-132">В этом примере вы позже назовет SHOWTASKPANE функцией, которая вызывает метод, а HIDETASKPANE — функцией, которая `Office.addin.showAsTaskpane` вызывает `Office.addin.hide` метод.</span><span class="sxs-lookup"><span data-stu-id="68c4b-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="68c4b-133">Создание сопоставления действий с их функциями</span><span class="sxs-lookup"><span data-stu-id="68c4b-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="68c4b-134">В проекте откройте файл JavaScript, загруженный вашей htmL-страницей в `<FunctionFile>` элементе.</span><span class="sxs-lookup"><span data-stu-id="68c4b-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="68c4b-135">В файле JavaScript [используйте API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотнося каждое действие, указанное в файле JSON, с функцией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="68c4b-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="68c4b-136">Добавьте в файл следующий JavaScript.</span><span class="sxs-lookup"><span data-stu-id="68c4b-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="68c4b-137">Обратите внимание на следующее о коде.</span><span class="sxs-lookup"><span data-stu-id="68c4b-137">Note the following about the code.</span></span>

    - <span data-ttu-id="68c4b-138">Первый параметр — это одно из действий из файла JSON.</span><span class="sxs-lookup"><span data-stu-id="68c4b-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="68c4b-139">Второй параметр — это функция, которая выполняется при нажатии клавиши на сочетание ключей, относясь к действию в файле JSON.</span><span class="sxs-lookup"><span data-stu-id="68c4b-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="68c4b-140">Чтобы продолжить пример, используйте `'SHOWTASKPANE'` в качестве первого параметра.</span><span class="sxs-lookup"><span data-stu-id="68c4b-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="68c4b-141">Для тела функции используйте [метод Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) для открытия области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="68c4b-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="68c4b-142">После этого код должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="68c4b-142">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="68c4b-143">Добавьте второй вызов функции, чтобы соединить действие с функцией, вызываемой `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide.](/javascript/api/office/office.addin#hide--)</span><span class="sxs-lookup"><span data-stu-id="68c4b-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="68c4b-144">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="68c4b-144">The following is an example.</span></span>

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

<span data-ttu-id="68c4b-145">Следуя предыдущим шагам, надстройка позволяет управлять видимостью области задач, нажимая **на Ctrl+Alt+Up** и **Ctrl+Alt+Down.**</span><span class="sxs-lookup"><span data-stu-id="68c4b-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Alt+Up** and **Ctrl+Alt+Down**.</span></span> <span data-ttu-id="68c4b-146">Такое же поведение показано в примере Excel клавиш в репо Office PnP надстройки в GitHub. [](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="68c4b-146">The same behavior is shown in the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample in the Office Add-ins PnP repo in GitHub.</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="68c4b-147">Сведения и ограничения</span><span class="sxs-lookup"><span data-stu-id="68c4b-147">Details and restrictions</span></span>

### <a name="construct-the-action-objects"></a><span data-ttu-id="68c4b-148">Построение объектов действия</span><span class="sxs-lookup"><span data-stu-id="68c4b-148">Construct the action objects</span></span>

<span data-ttu-id="68c4b-149">Используйте следующие рекомендации при указании объектов в массиве `actions` shortcuts.js.</span><span class="sxs-lookup"><span data-stu-id="68c4b-149">Use the following guidelines when specifying the objects in the `actions` array of the shortcuts.json.</span></span>

- <span data-ttu-id="68c4b-150">Имена свойств `id` и `name` обязательны.</span><span class="sxs-lookup"><span data-stu-id="68c4b-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="68c4b-151">Свойство `id` используется для уникальной идентификации действия, вызываемого с помощью ярлыка клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68c4b-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="68c4b-152">Свойство `name` должно быть удобной строкой, описываемой действием.</span><span class="sxs-lookup"><span data-stu-id="68c4b-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="68c4b-153">Это должно быть сочетание символов A - Z, a - z, 0 - 9 и знаков препинания "-", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="68c4b-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="68c4b-154">Свойство `type`— необязательное.</span><span class="sxs-lookup"><span data-stu-id="68c4b-154">The `type` property is optional.</span></span> <span data-ttu-id="68c4b-155">В `ExecuteFunction` настоящее время поддерживается только тип.</span><span class="sxs-lookup"><span data-stu-id="68c4b-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="68c4b-156">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="68c4b-156">The following is an example.</span></span>

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

<span data-ttu-id="68c4b-157">Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="68c4b-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="construct-the-shortcut-objects"></a><span data-ttu-id="68c4b-158">Построение объектов ярлыка</span><span class="sxs-lookup"><span data-stu-id="68c4b-158">Construct the shortcut objects</span></span>

<span data-ttu-id="68c4b-159">Используйте следующие рекомендации при указании объектов в массиве `shortcuts` shortcuts.js.</span><span class="sxs-lookup"><span data-stu-id="68c4b-159">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json.</span></span>

- <span data-ttu-id="68c4b-160">Имена свойств `action` `key` и `default` обязательно.</span><span class="sxs-lookup"><span data-stu-id="68c4b-160">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="68c4b-161">Значение свойства является строкой и должно соответствовать одному из свойств `action` `id` объекта действия.</span><span class="sxs-lookup"><span data-stu-id="68c4b-161">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="68c4b-162">Свойство может быть любым сочетанием символов A - Z, a -z, 0 - 9, а знаки препинания `default` "-", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="68c4b-162">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="68c4b-163">(По соглашению в этих свойствах не используются буквы более низкого уровня.)</span><span class="sxs-lookup"><span data-stu-id="68c4b-163">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="68c4b-164">Свойство должно содержать имя по крайней мере одного ключа модификатора `default` (Alt, Ctrl, Shift) и только одного другого ключа.</span><span class="sxs-lookup"><span data-stu-id="68c4b-164">The `default` property must contain the name of at least one modifier key (Alt, Ctrl, Shift) and only one other key.</span></span> 
- <span data-ttu-id="68c4b-165">Shift не может использоваться в качестве только ключа модификатора.</span><span class="sxs-lookup"><span data-stu-id="68c4b-165">Shift cannot be used as the only modifier key.</span></span> <span data-ttu-id="68c4b-166">Объединяйте Shift с Alt или Ctrl.</span><span class="sxs-lookup"><span data-stu-id="68c4b-166">Combine Shift with either Alt or Ctrl.</span></span>
- <span data-ttu-id="68c4b-167">Для macs мы также поддерживаем ключ модификатора Команд.</span><span class="sxs-lookup"><span data-stu-id="68c4b-167">For Macs, we also support the Command modifier key.</span></span>
- <span data-ttu-id="68c4b-168">Для macs Alt соо-</span><span class="sxs-lookup"><span data-stu-id="68c4b-168">For Macs, Alt is mapped to the Option key.</span></span> <span data-ttu-id="68c4b-169">Для Windows командной командой нажата клавиша Ctrl.</span><span class="sxs-lookup"><span data-stu-id="68c4b-169">For Windows, Command is mapped to the Ctrl key.</span></span>
- <span data-ttu-id="68c4b-170">Если два символа связаны с одним и тем же физическим ключом в стандартной клавиатуре, они являются синонимами в свойстве; например, Alt+a и Alt+A являются одним и тем же ярлыком, как `default` и Ctrl+- и Ctrl+, так как "-" и "_" являются одним и тем же физическим \_ ключом.</span><span class="sxs-lookup"><span data-stu-id="68c4b-170">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, Alt+a and Alt+A are the same shortcut, so are Ctrl+- and Ctrl+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="68c4b-171">Символ "+" указывает, что клавиши с обеих сторон нажаты одновременно.</span><span class="sxs-lookup"><span data-stu-id="68c4b-171">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="68c4b-172">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="68c4b-172">The following is an example.</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

<span data-ttu-id="68c4b-173">Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="68c4b-173">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="68c4b-174">Клавиши KeyTips, также известные как последовательное клавиши, такие как ярлык Excel для выбора цвета заполнения **Alt+H, H,** не поддерживаются в Office надстроек.</span><span class="sxs-lookup"><span data-stu-id="68c4b-174">KeyTips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a><span data-ttu-id="68c4b-175">Избегайте комбинаций ключей, которые используются другими надстройки</span><span class="sxs-lookup"><span data-stu-id="68c4b-175">Avoid key combinations in use by other add-ins</span></span>

<span data-ttu-id="68c4b-176">Существует множество клавиш, которые уже используются Office.</span><span class="sxs-lookup"><span data-stu-id="68c4b-176">There are many keyboard shortcuts that are already in use by Office.</span></span> <span data-ttu-id="68c4b-177">Избегайте регистрации клавишных ярлыков для надстройки, которые уже используются, однако могут существовать некоторые случаи, когда необходимо переопределять существующие ярлыки клавиатуры или обрабатывать конфликты между несколькими надстройки, которые зарегистрировали один и тот же ярлык клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68c4b-177">Avoid registering keyboard shortcuts for your add-in that are already in use, however there may be some instances where it is necessary to override existing keyboard shortcuts or handle conflicts between multiple add-ins that have registered the same keyboard shortcut.</span></span>

<span data-ttu-id="68c4b-178">В случае конфликта пользователь увидит диалоговое окно при первой попытке использовать конфликтующий ярлык клавиатуры, обратите внимание, что имя действия, отображаемого в этом диалоговом диалоговом окне, является свойством в объекте действия в `name` `shortcuts.json` файле.</span><span class="sxs-lookup"><span data-stu-id="68c4b-178">In the case of a conflict, the user will see a dialog box the first time they attempt to use a conflicting keyboard shortcut, note that the action name that is displayed in this dialog is the `name` property in the action object in `shortcuts.json` file.</span></span>

![Иллюстрация, показывающая конфликтный модал с двумя разными действиями для одного ярлыка.](../images/add-in-shortcut-conflict-modal.png)

<span data-ttu-id="68c4b-180">Пользователь может выбрать, какое действие будет принимать ярлык клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68c4b-180">The user can select which action the keyboard shortcut will take.</span></span> <span data-ttu-id="68c4b-181">После выбора предпочтения сохраняются для будущих применений одного и того же ярлыка.</span><span class="sxs-lookup"><span data-stu-id="68c4b-181">After making the selection, the preference is saved for future uses of the same shortcut.</span></span> <span data-ttu-id="68c4b-182">Параметры ярлыка сохраняются для каждого пользователя, для платформы.</span><span class="sxs-lookup"><span data-stu-id="68c4b-182">The shortcut preferences are saved per user, per platform.</span></span> <span data-ttu-id="68c4b-183">Если пользователь хочет изменить свои предпочтения, он может **вызвать команду быстрого** доступа Office надстройки из поискового окна Tell **me.**</span><span class="sxs-lookup"><span data-stu-id="68c4b-183">If the user wishes to change their preferences, they can invoke the **Reset Office Add-ins shortcut preferences** command from the **Tell me** search box.</span></span> <span data-ttu-id="68c4b-184">При наводке команда очищает все параметры ярлыка надстройки пользователя, и пользователю снова будет предложен диалоговое окно конфликта при следующей попытке использовать конфликтующий ярлык:</span><span class="sxs-lookup"><span data-stu-id="68c4b-184">Invoking the command clears all of the user's add-in shortcut preferences and the user will again be prompted with the conflict dialog box the next time they attempt to use a conflicting shortcut:</span></span>

![Поле поиска Tell me в Excel с указанием действия Office настройки ярлыка надстройки.](../images/add-in-reset-shortcuts-action.png)

<span data-ttu-id="68c4b-186">Для наилучшего пользовательского интерфейса рекомендуется свести к минимуму конфликты с Excel с помощью этих методов:</span><span class="sxs-lookup"><span data-stu-id="68c4b-186">For the best user experience, we recommend that you minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="68c4b-187">Используйте только клавиши со следующим шаблоном: \**Ctrl+Shift+Alt+* x\*\*\*\*, где *x* — это другой ключ.</span><span class="sxs-lookup"><span data-stu-id="68c4b-187">Use only keyboard shortcuts with the following pattern: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="68c4b-188">Если вам нужно больше клавиш, ознакомьтесь со списком Excel клавиш [и](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)не применяйте их в надстройки.</span><span class="sxs-lookup"><span data-stu-id="68c4b-188">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>
- <span data-ttu-id="68c4b-189">Когда фокус клавиатуры находится внутри пользовательского интерфейса надстройки, **Ctrl+Spacebar** и **Ctrl+Shift+F10** не будут работать, так как это основные ярлыки доступности.</span><span class="sxs-lookup"><span data-stu-id="68c4b-189">When the keyboard focus is inside the add-in UI, **Ctrl+Spacebar** and **Ctrl+Shift+F10** will not work as these are essential accessibility shortcuts.</span></span>
- <span data-ttu-id="68c4b-190">На компьютере Windows или Mac, если в меню поиска недоступна команда "Reset Office надстройки", пользователь может вручную добавить команду в ленту, настроив ленту через контекстное меню.</span><span class="sxs-lookup"><span data-stu-id="68c4b-190">On a Windows or Mac computer, if the "Reset Office Add-ins shortcut preferences" command is not available on the search menu, the user can manually add the command to the ribbon by customizing the ribbon through the context menu.</span></span>

## <a name="customize-the-keyboard-shortcuts-per-platform"></a><span data-ttu-id="68c4b-191">Настройка ярлыков клавиатуры для платформы</span><span class="sxs-lookup"><span data-stu-id="68c4b-191">Customize the keyboard shortcuts per platform</span></span>

<span data-ttu-id="68c4b-192">Можно настроить ярлыки для конкретной платформы.</span><span class="sxs-lookup"><span data-stu-id="68c4b-192">It's possible to customize shortcuts to be platform-specific.</span></span> <span data-ttu-id="68c4b-193">Ниже приводится пример объекта, который настраивает ярлыки для каждой из следующих `shortcuts` платформ: `windows` , , `mac` `web` .</span><span class="sxs-lookup"><span data-stu-id="68c4b-193">The following is an example of the `shortcuts` object that customizes the shortcuts for each of the following platforms: `windows`, `mac`, `web`.</span></span> <span data-ttu-id="68c4b-194">Обратите внимание, что для каждого ярлыка необходимо иметь клавишу `default` ярлыка.</span><span class="sxs-lookup"><span data-stu-id="68c4b-194">Note that you must still have a `default` shortcut key for each shortcut.</span></span>

<span data-ttu-id="68c4b-195">В следующем примере `default` ключом является клавиша отката для любой платформы, которая не указана.</span><span class="sxs-lookup"><span data-stu-id="68c4b-195">In the following example, the `default` key is the fallback key for any platform that is not specified.</span></span> <span data-ttu-id="68c4b-196">Единственная не указанная платформа Windows, поэтому ключ будет применяться только к `default` Windows.</span><span class="sxs-lookup"><span data-stu-id="68c4b-196">The only platform not specified is Windows, so the `default` key will only apply to Windows.</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="68c4b-197">Локализовать ярлыки клавиатуры JSON</span><span class="sxs-lookup"><span data-stu-id="68c4b-197">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="68c4b-198">Если надстройка поддерживает несколько локалов, необходимо локализовать свойство `name` объектов действия.</span><span class="sxs-lookup"><span data-stu-id="68c4b-198">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="68c4b-199">Кроме того, если в любом из локаутах, поддерживаюх надстройку, есть алфавиты или различные системы записи, а значит, и другие клавиатуры, возможно, потребуется также локализовать ярлыки.</span><span class="sxs-lookup"><span data-stu-id="68c4b-199">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="68c4b-200">Сведения о том, как локализовать клавиши ярлыков JSON, см. в рубрезе [Localize extended overrides.](../develop/localization.md#localize-extended-overrides)</span><span class="sxs-lookup"><span data-stu-id="68c4b-200">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="68c4b-201">Ярлыки браузера, которые нельзя переопределять</span><span class="sxs-lookup"><span data-stu-id="68c4b-201">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="68c4b-202">При использовании настраиваемого сочетания клавиш в Интернете некоторые клавиши, используемые браузером, не могут быть переопределены надстройки. Этот список находится в процессе выполнения.</span><span class="sxs-lookup"><span data-stu-id="68c4b-202">When using custom keyboard shortcuts on the web, some keyboard shortcuts that are used by the browser cannot be overridden by add-ins. This list is a work in progress.</span></span> <span data-ttu-id="68c4b-203">Если вы обнаружите другие комбинации, которые нельзя переопределять, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.</span><span class="sxs-lookup"><span data-stu-id="68c4b-203">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="68c4b-204">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="68c4b-204">Ctrl+N</span></span>
- <span data-ttu-id="68c4b-205">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="68c4b-205">Ctrl+Shift+N</span></span>
- <span data-ttu-id="68c4b-206">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="68c4b-206">Ctrl+T</span></span>
- <span data-ttu-id="68c4b-207">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="68c4b-207">Ctrl+Shift+T</span></span>
- <span data-ttu-id="68c4b-208">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="68c4b-208">Ctrl+W</span></span>
- <span data-ttu-id="68c4b-209">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="68c4b-209">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="68c4b-210">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="68c4b-210">Next Steps</span></span>

- <span data-ttu-id="68c4b-211">См. [Excel надстройки](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) для клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68c4b-211">See the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample add-in.</span></span>
- <span data-ttu-id="68c4b-212">Получите обзор работы с расширенными переопределениями в Работе с расширенными [переопределениями манифеста.](../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="68c4b-212">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>
