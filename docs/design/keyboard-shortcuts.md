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
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="95950-103">Добавление настраиваемых сочетаний клавиш в надстройки Office (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="95950-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="95950-104">Сочетания клавиш, называемые также сочетаниями клавиш, позволяют пользователям вашей надстройки работать эффективнее и расширять возможности надстройки для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.</span><span class="sxs-lookup"><span data-stu-id="95950-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="95950-105">Чтобы начать работу с рабочей версией надстройки с включенными сочетаниями клавиш, выполните клонирование и выполните примеры сочетаний [клавиш Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span><span class="sxs-lookup"><span data-stu-id="95950-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="95950-106">Когда вы будете готовы добавить сочетания клавиш для своей надстройки, перейдите к этой статье.</span><span class="sxs-lookup"><span data-stu-id="95950-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="95950-107">Добавление сочетаний клавиш в надстройку состоит из трех этапов:</span><span class="sxs-lookup"><span data-stu-id="95950-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="95950-108">[Настройте манифест надстройки](#configure-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="95950-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="95950-109">[Создайте или измените файл ярлыков JSON](#create-or-edit-the-shortcuts-json-file) , чтобы определить действия и их сочетания клавиш.</span><span class="sxs-lookup"><span data-stu-id="95950-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="95950-110">[Добавьте один или несколько вызовов среды выполнения](#create-a-mapping-of-actions-to-their-functions) API [Office. Actions.](/javascript/api/office/office.actions#associate) Map, чтобы сопоставить функцию с каждым действием.</span><span class="sxs-lookup"><span data-stu-id="95950-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="95950-111">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="95950-111">Configure the manifest</span></span>

<span data-ttu-id="95950-112">В манифесте есть два небольших изменения, которые необходимо выполнить.</span><span class="sxs-lookup"><span data-stu-id="95950-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="95950-113">Один — позволить надстройке использовать общую среду выполнения, а другая — указать на файл в формате JSON, в котором были определены сочетания клавиш.</span><span class="sxs-lookup"><span data-stu-id="95950-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="95950-114">Настройка надстройки для использования общей среды выполнения</span><span class="sxs-lookup"><span data-stu-id="95950-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="95950-115">Добавление пользовательских сочетаний клавиш требует, чтобы ваша надстройка использовала общую среду выполнения.</span><span class="sxs-lookup"><span data-stu-id="95950-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="95950-116">Для получения дополнительных сведений [Настройте надстройку для использования общей среды выполнения](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="95950-116">For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="95950-117">Связывание файла сопоставления с манифестом</span><span class="sxs-lookup"><span data-stu-id="95950-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="95950-118">Непосредственно *ниже* (не внутри) `<VersionOverrides>` элемента в манифесте добавьте элемент [екстендедоверридес](../reference/manifest/extendedoverrides.md) .</span><span class="sxs-lookup"><span data-stu-id="95950-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="95950-119">Присвойте `Url` атрибуту полный URL-адрес JSON-файла в проекте, который будет создан на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="95950-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="95950-120">Создание или редактирование файла ярлыков JSON</span><span class="sxs-lookup"><span data-stu-id="95950-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="95950-121">Создайте файл JSON в проекте.</span><span class="sxs-lookup"><span data-stu-id="95950-121">Create a JSON file in your project.</span></span> <span data-ttu-id="95950-122">Убедитесь, что путь к файлу совпадает с расположением, указанным для `Url` атрибута элемента [екстендедоверридес](../reference/manifest/extendedoverrides.md) .</span><span class="sxs-lookup"><span data-stu-id="95950-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="95950-123">В этом файле будут описаны сочетания клавиш и действия, которые они будут вызывать.</span><span class="sxs-lookup"><span data-stu-id="95950-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="95950-124">В файле JSON существует два массива.</span><span class="sxs-lookup"><span data-stu-id="95950-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="95950-125">Массив Actions будет содержать объекты, определяющие действия, которые необходимо вызвать, а массив ярлыков будет содержать объекты, которые сопоставлены с сочетаниями клавиш на действия.</span><span class="sxs-lookup"><span data-stu-id="95950-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="95950-126">Вот пример:</span><span class="sxs-lookup"><span data-stu-id="95950-126">Here is an example:</span></span>

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

    <span data-ttu-id="95950-127">Дополнительные сведения об объектах JSON приведены в статье [Создание объектов Action](#constructing-the-action-objects) и [Создание объектов ярлыков](#constructing-the-shortcut-objects).</span><span class="sxs-lookup"><span data-stu-id="95950-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="95950-128">Полная схема для ярлыков JSON [extended-manifest.schema.jsвключена](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="95950-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="95950-129">(Примечание: ссылка на схему может не работать рано в периоде предварительной версии.)</span><span class="sxs-lookup"><span data-stu-id="95950-129">(Note: The link to the schema may not be working early in the preview period.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="95950-130">В этой статье можно использовать элемент управления вместо "CTRL".</span><span class="sxs-lookup"><span data-stu-id="95950-130">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="95950-131">На более позднем этапе действия будут сопоставлены с написанными функциями.</span><span class="sxs-lookup"><span data-stu-id="95950-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="95950-132">В этом примере позднее показано сопоставление SHOWTASKPANE с функцией, которая вызывает `Office.addin.showAsTaskpane` метод, и хидетаскпане в функцию, которая вызывает `Office.addin.hide` метод.</span><span class="sxs-lookup"><span data-stu-id="95950-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="95950-133">Создание сопоставления действий с их функциями</span><span class="sxs-lookup"><span data-stu-id="95950-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="95950-134">В проекте откройте файл JavaScript, загруженный HTML-страницей в `<FunctionFile>` элементе.</span><span class="sxs-lookup"><span data-stu-id="95950-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="95950-135">В файле JavaScript используйте API [Office. Actions.](/javascript/api/office/office.actions#associate) Map, чтобы сопоставить каждое действие, указанное в JSON-файле, с функцией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="95950-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="95950-136">Добавьте в файл приведенный ниже код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="95950-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="95950-137">Обратите внимание на следующие особенности кода:</span><span class="sxs-lookup"><span data-stu-id="95950-137">Note the following about the code:</span></span>

    - <span data-ttu-id="95950-138">Первый параметр — это одно из действий из JSON-файла.</span><span class="sxs-lookup"><span data-stu-id="95950-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="95950-139">Второй параметр — это функция, которая запускается, когда пользователь нажимает комбинацию клавиш, сопоставленную с действием в JSON-файле.</span><span class="sxs-lookup"><span data-stu-id="95950-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="95950-140">Чтобы продолжить пример, используйте `'SHOWTASKPANE'` в качестве первого параметра.</span><span class="sxs-lookup"><span data-stu-id="95950-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="95950-141">Для основной части функции используйте метод [Office. AddIn. showTaskpane](/javascript/api/office/office.addin#showastaskpane--) , чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="95950-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="95950-142">После завершения код должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="95950-142">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="95950-143">Добавьте второй вызов `Office.actions.associate` функции, чтобы сопоставить `HIDETASKPANE` действие с функцией, которая вызывает [Office. AddIn. Hide](/javascript/api/office/office.addin#hide--).</span><span class="sxs-lookup"><span data-stu-id="95950-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="95950-144">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="95950-144">The following is an example:</span></span>

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

<span data-ttu-id="95950-145">После выполнения описанных выше действий надстройка позволяет переключать видимость области задач, нажимая клавиши **Ctrl + Shift + стрелка вверх** и **Ctrl + Shift + стрелка вниз**.</span><span class="sxs-lookup"><span data-stu-id="95950-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="95950-146">Это то же поведение, которое показано в [примере надстройки "сочетания клавиш Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)".</span><span class="sxs-lookup"><span data-stu-id="95950-146">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="95950-147">Сведения и ограничения</span><span class="sxs-lookup"><span data-stu-id="95950-147">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="95950-148">Создание объектов Action</span><span class="sxs-lookup"><span data-stu-id="95950-148">Constructing the action objects</span></span>

<span data-ttu-id="95950-149">При указании объектов в массиве shortcuts.jsследует придерживаться следующих рекомендаций `action` .</span><span class="sxs-lookup"><span data-stu-id="95950-149">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="95950-150">Имена свойств `id` и `name` являются обязательными.</span><span class="sxs-lookup"><span data-stu-id="95950-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="95950-151">`id`Свойство используется для уникальной идентификации действия, которое вызывается с помощью сочетания клавиш.</span><span class="sxs-lookup"><span data-stu-id="95950-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="95950-152">`name`Свойство должно представлять собой удобную пользователю строку, описывающую действие.</span><span class="sxs-lookup"><span data-stu-id="95950-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="95950-153">Он должен быть комбинацией символов A – Z, a – z, 0-9 и знаков препинания "–", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="95950-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="95950-154">Свойство `type`— необязательное.</span><span class="sxs-lookup"><span data-stu-id="95950-154">The `type` property is optional.</span></span> <span data-ttu-id="95950-155">В настоящее время `ExecuteFunction` поддерживается только тип.</span><span class="sxs-lookup"><span data-stu-id="95950-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="95950-156">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="95950-156">The following is an example:</span></span>

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

<span data-ttu-id="95950-157">Полная схема для ярлыков JSON [extended-manifest.schema.jsвключена](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="95950-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="95950-158">(Примечание: ссылка на схему может не работать рано в периоде предварительной версии.)</span><span class="sxs-lookup"><span data-stu-id="95950-158">(Note: The link to the schema may not be working early in the preview period.)</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="95950-159">Создание объектов ярлыков</span><span class="sxs-lookup"><span data-stu-id="95950-159">Constructing the shortcut objects</span></span>

<span data-ttu-id="95950-160">При указании объектов в массиве shortcuts.jsследует придерживаться следующих рекомендаций `shortcuts` .</span><span class="sxs-lookup"><span data-stu-id="95950-160">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="95950-161">Имена свойств `action` `key` и `default` обязательные.</span><span class="sxs-lookup"><span data-stu-id="95950-161">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="95950-162">Значение `action` свойства является строкой и должно удовлетворять одному из `id` свойств в объекте Action.</span><span class="sxs-lookup"><span data-stu-id="95950-162">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="95950-163">`default`Свойство может быть любым сочетанием символов a – z, a – z, 0-9 и знаков препинания "–", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="95950-163">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="95950-164">(В соответствии с соглашением буквы нижнего регистра не используются в этих свойствах.)</span><span class="sxs-lookup"><span data-stu-id="95950-164">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="95950-165">`default`Свойство должно содержать имя по крайней мере одной клавиши-модификатора (Alt, CTRL, Shift) и только один ключ.</span><span class="sxs-lookup"><span data-stu-id="95950-165">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="95950-166">Для Макинтош мы также поддерживаем клавишей CTRL COMMAND.</span><span class="sxs-lookup"><span data-stu-id="95950-166">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="95950-167">Для Макинтошей атрибут ALT сопоставлен с ключом OPTION.</span><span class="sxs-lookup"><span data-stu-id="95950-167">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="95950-168">Для Windows команда сопоставляется с клавишей CTRL.</span><span class="sxs-lookup"><span data-stu-id="95950-168">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="95950-169">Если два символа связаны с одним и тем же физическим ключом на стандартной клавиатуре, то они являются синонимами в `default` свойстве, например ALT + a, а Alt + a — это одно сочетание клавиш, поэтому клавиши CTRL +-и CTRL +, \_ так как "-" и "_" являются одним и тем же физическим ключом.</span><span class="sxs-lookup"><span data-stu-id="95950-169">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="95950-170">Символ "+" указывает на то, что клавиши с любой стороны объекта одновременно нажаты.</span><span class="sxs-lookup"><span data-stu-id="95950-170">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="95950-171">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="95950-171">The following is an example:</span></span>

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

<span data-ttu-id="95950-172">Полная схема для ярлыков JSON [extended-manifest.schema.jsвключена](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="95950-172">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="95950-173">(Примечание: ссылка на схему может не работать рано в периоде предварительной версии.)</span><span class="sxs-lookup"><span data-stu-id="95950-173">(Note: The link to the schema may not be working early in the preview period.)</span></span>

> [!NOTE]
> <span data-ttu-id="95950-174">Подсказки, которые также называются последовательной клавишей, такие как ярлык Excel для выбора цвета заливки **ALT + H, H**, не поддерживаются в надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="95950-174">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="95950-175">Использование сочетаний клавиш, когда фокус находится в области задач</span><span class="sxs-lookup"><span data-stu-id="95950-175">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="95950-176">В настоящее время сочетания клавиш для надстройки Office могут вызываться только в том случае, если фокус пользователя находится на листе.</span><span class="sxs-lookup"><span data-stu-id="95950-176">Currently, the keyboard shortcuts for an Office add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="95950-177">Когда фокус пользователя находится в пользовательском интерфейсе Office (например, область задач), ни одна из ее ярлыков не игнорируется.</span><span class="sxs-lookup"><span data-stu-id="95950-177">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="95950-178">В качестве обходного решения надстройка может определить обработчики клавиатуры, которые могут вызывать определенные действия, когда фокус пользователя находится в пользовательском интерфейсе надстройки.</span><span class="sxs-lookup"><span data-stu-id="95950-178">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="95950-179">Использование сочетаний клавиш, которые уже используются в Office или другой надстройке</span><span class="sxs-lookup"><span data-stu-id="95950-179">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="95950-180">В течение периода предварительного просмотра нет системы для определения действий, которые происходят, когда пользователь нажимает сочетание клавиш, зарегистрированное надстройкой, а также Office или другой надстройкой.</span><span class="sxs-lookup"><span data-stu-id="95950-180">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="95950-181">Поведение не определено.</span><span class="sxs-lookup"><span data-stu-id="95950-181">Behavior is undefined.</span></span>

<span data-ttu-id="95950-182">В настоящее время не существует решения, в котором две или более надстройки зарегистрировали одну комбинацию клавиш, но вы можете минимизировать конфликты с Excel, выполнив приведенные ниже рекомендации.</span><span class="sxs-lookup"><span data-stu-id="95950-182">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="95950-183">Используйте только сочетания клавиш со следующим шаблоном в надстройке: \**CTRL + SHIFT + ALT +* x \* \* \*, где *x* — это другой ключ.</span><span class="sxs-lookup"><span data-stu-id="95950-183">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="95950-184">Если вам нужны дополнительные сочетания клавиш, проверьте список сочетаний [клавиш Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)и старайтесь не использовать их в вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="95950-184">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="95950-185">Ярлыки браузеров, которые не могут быть переопределены</span><span class="sxs-lookup"><span data-stu-id="95950-185">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="95950-186">Вы не можете использовать следующие сочетания клавиш.</span><span class="sxs-lookup"><span data-stu-id="95950-186">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="95950-187">Они используются браузерами и не могут быть переопределены.</span><span class="sxs-lookup"><span data-stu-id="95950-187">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="95950-188">Этот список является рабочим процессом.</span><span class="sxs-lookup"><span data-stu-id="95950-188">This list is a work in progress.</span></span> <span data-ttu-id="95950-189">Если вы обнаружите другие сочетания, которые невозможно переопределить, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.</span><span class="sxs-lookup"><span data-stu-id="95950-189">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="95950-190">Ctrl + N</span><span class="sxs-lookup"><span data-stu-id="95950-190">Ctrl+N</span></span>
- <span data-ttu-id="95950-191">Ctrl + Shift + N</span><span class="sxs-lookup"><span data-stu-id="95950-191">Ctrl+Shift+N</span></span>
- <span data-ttu-id="95950-192">Ctrl + T</span><span class="sxs-lookup"><span data-stu-id="95950-192">Ctrl+T</span></span>
- <span data-ttu-id="95950-193">Ctrl + Shift + T</span><span class="sxs-lookup"><span data-stu-id="95950-193">Ctrl+Shift+T</span></span>
- <span data-ttu-id="95950-194">Ctrl + W</span><span class="sxs-lookup"><span data-stu-id="95950-194">Ctrl+W</span></span>
- <span data-ttu-id="95950-195">Ctrl + ПГУП/Пгдн</span><span class="sxs-lookup"><span data-stu-id="95950-195">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="95950-196">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="95950-196">Next Steps</span></span>

- <span data-ttu-id="95950-197">В этой статье приведены примеры сочетаний [клавиш Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)для надстроек.</span><span class="sxs-lookup"><span data-stu-id="95950-197">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
