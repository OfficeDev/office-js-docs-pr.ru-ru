---
title: Вставка и удаление слайдов в презентации PowerPoint
description: Сведения о том, как вставлять слайды из одной презентации в другую и удалять слайды.
ms.date: 12/04/2020
localization_priority: Normal
ms.openlocfilehash: ceb78054a95ac4b26bd71f79a086a00e3dce5278
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/09/2020
ms.locfileid: "49613711"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a><span data-ttu-id="4624c-103">Вставка и удаление слайдов в презентации PowerPoint (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="4624c-103">Insert and delete slides in a PowerPoint presentation (preview)</span></span>

<span data-ttu-id="4624c-104">Надстройка PowerPoint позволяет вставлять слайды из одной презентации в текущую, используя библиотеку JavaScript, зависящую от приложения PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="4624c-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="4624c-105">Вы можете указать, следует ли вставлять вставленные слайды в исходную презентацию или форматировать целевую презентацию.</span><span class="sxs-lookup"><span data-stu-id="4624c-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="4624c-106">Вы также можете удалять слайды из презентации.</span><span class="sxs-lookup"><span data-stu-id="4624c-106">You can also delete slides from the presentation.</span></span>

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

<span data-ttu-id="4624c-107">API вставки слайдов в основном используются в сценариях презентации: существует небольшое количество известных презентаций, которые могут быть вставлены надстройкой в виде пулов слайдов.</span><span class="sxs-lookup"><span data-stu-id="4624c-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="4624c-108">В этом сценарии либо вы, либо клиент должны создать и поддерживать источник данных, который соответствует условию выбора (например, заголовки слайдов или изображения) с идентификаторами слайдов.</span><span class="sxs-lookup"><span data-stu-id="4624c-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="4624c-109">API также можно использовать в сценариях, где пользователь может вставлять слайды из произвольной презентации, но в этом сценарии пользователь практически ограничен вставкой *всех* слайдов из исходной презентации.</span><span class="sxs-lookup"><span data-stu-id="4624c-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="4624c-110">Дополнительные сведения об этом [можно узнать в разделе Выбор слайдов для вставки](#selecting-which-slides-to-insert) .</span><span class="sxs-lookup"><span data-stu-id="4624c-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="4624c-111">Вставить слайды из одной презентации в другую можно двумя шагами.</span><span class="sxs-lookup"><span data-stu-id="4624c-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="4624c-112">Преобразование исходного файла презентации (PPTX) в строку в формате Base64.</span><span class="sxs-lookup"><span data-stu-id="4624c-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="4624c-113">Используйте `insertSlidesFromBase64` метод, чтобы вставить один или несколько слайдов из файла Base64 в текущую презентацию.</span><span class="sxs-lookup"><span data-stu-id="4624c-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="4624c-114">Преобразование исходной презентации в формат Base64</span><span class="sxs-lookup"><span data-stu-id="4624c-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="4624c-115">Существует множество способов преобразования файла в Base64.</span><span class="sxs-lookup"><span data-stu-id="4624c-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="4624c-116">Выбор используемого языка программирования и библиотеки, а также необходимость преобразования на стороне сервера надстройки или на стороне клиента определяется сценарием.</span><span class="sxs-lookup"><span data-stu-id="4624c-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="4624c-117">Чаще всего преобразование в JavaScript на стороне клиента выполняется с помощью объекта [FileReader браузером](https://developer.mozilla.org/docs/Web/API/FileReader) .</span><span class="sxs-lookup"><span data-stu-id="4624c-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="4624c-118">В приведенном ниже примере показана эта практика.</span><span class="sxs-lookup"><span data-stu-id="4624c-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="4624c-119">Сначала получите ссылку на исходный файл PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="4624c-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="4624c-120">В этом примере мы будем использовать `<input>` элемент управления типа, `file` чтобы предлагать пользователю выбрать файл.</span><span class="sxs-lookup"><span data-stu-id="4624c-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="4624c-121">Добавьте указанную ниже разметку на страницу надстройки.</span><span class="sxs-lookup"><span data-stu-id="4624c-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="4624c-122">Эта разметка добавляет пользовательский интерфейс на страницу на следующем снимке экрана:</span><span class="sxs-lookup"><span data-stu-id="4624c-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![Снимок экрана с элементом управления вводом типа HTML-файла, которому предшествует пояснительное предложение чтение "Выбор презентации PowerPoint, из которой нужно вставить слайды".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="4624c-125">Существует множество других способов получения файла PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="4624c-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="4624c-126">Например, если файл хранится в OneDrive или SharePoint, вы можете скачать его с помощью Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="4624c-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="4624c-127">Дополнительные сведения см. [в статье работа с файлами в Microsoft Graph](/graph/api/resources/onedrive) и [доступ к файлам с помощью Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span><span class="sxs-lookup"><span data-stu-id="4624c-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="4624c-128">Добавьте следующий код в JavaScript надстройки, чтобы назначить функцию для события элемента управления вводом `change` .</span><span class="sxs-lookup"><span data-stu-id="4624c-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="4624c-129">(Вы создадите `storeFileAsBase64` функцию на следующем шаге.)</span><span class="sxs-lookup"><span data-stu-id="4624c-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="4624c-130">Добавьте в него указанный ниже код.</span><span class="sxs-lookup"><span data-stu-id="4624c-130">Add the following code.</span></span> <span data-ttu-id="4624c-131">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="4624c-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="4624c-132">`reader.readAsDataURL`Метод преобразует файл в формат Base64 и сохраняет его в `reader.result` свойстве.</span><span class="sxs-lookup"><span data-stu-id="4624c-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="4624c-133">После выполнения метода он запускает `onload` обработчик событий.</span><span class="sxs-lookup"><span data-stu-id="4624c-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="4624c-134">`onload`Обработчик событий удаляет метаданные из зашифрованного файла и сохраняет закодированную строку в глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="4624c-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="4624c-135">Строка в кодировке Base64 хранится глобально, так как она будет прочитана другой функцией, созданной на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="4624c-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="4624c-136">Вставка слайдов с помощью insertSlidesFromBase64</span><span class="sxs-lookup"><span data-stu-id="4624c-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="4624c-137">Надстройка вставляет слайды из другой презентации PowerPoint в текущую презентацию с помощью метода [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) .</span><span class="sxs-lookup"><span data-stu-id="4624c-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="4624c-138">Ниже приведен простой пример, в котором все слайды исходной презентации вставляются в начало текущей презентации, а вставленные слайды сохраняются в формате исходного файла.</span><span class="sxs-lookup"><span data-stu-id="4624c-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="4624c-139">Обратите внимание, что `chosenFileBase64` это глобальная переменная, содержащая версию файла презентации PowerPoint в кодировке Base64.</span><span class="sxs-lookup"><span data-stu-id="4624c-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="4624c-140">Вы можете управлять некоторыми аспектами вставки, включая место вставки слайдов и способ получения исходного или целевого форматирования, путем передачи объекта [инсертслидеоптионс](/javascript/api/powerpoint/powerpoint.insertslideoptions) в качестве второго параметра `insertSlidesFromBase64` .</span><span class="sxs-lookup"><span data-stu-id="4624c-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="4624c-141">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4624c-141">The following is an example.</span></span> <span data-ttu-id="4624c-142">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="4624c-142">About this code, note:</span></span>

- <span data-ttu-id="4624c-143">Свойство имеет два возможных значения `formatting` : "уседестинатионсеме" и "кипсаурцеформаттинг".</span><span class="sxs-lookup"><span data-stu-id="4624c-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="4624c-144">При необходимости можно использовать `InsertSlideFormatting` Перечисление (например, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).</span><span class="sxs-lookup"><span data-stu-id="4624c-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="4624c-145">Функция вставит слайды из исходной презентации сразу же после слайда, указанного `targetSlideId` свойством.</span><span class="sxs-lookup"><span data-stu-id="4624c-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="4624c-146">Значение этого свойства является строкой одной из трех возможных форм: \***nnn \* #**, \* *#* ммммммммм \* \* \* или \**_nnn_ #* ммммммммм \* \* \*, где *nnn* — идентификатор слайда (обычно 3 цифры), а *ммммммммм* — идентификатор создания слайда (обычно 9 цифры).</span><span class="sxs-lookup"><span data-stu-id="4624c-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="4624c-147">Некоторые примеры: `267#763315295` , `267#` и `#763315295` .</span><span class="sxs-lookup"><span data-stu-id="4624c-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

<span data-ttu-id="4624c-148">Конечно, вы, как правило, не узнаете на момент кодирования идентификатор или идентификатор создания целевого слайда.</span><span class="sxs-lookup"><span data-stu-id="4624c-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="4624c-149">Чаще всего надстройка запрашивает у пользователей Выбор целевого слайда.</span><span class="sxs-lookup"><span data-stu-id="4624c-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="4624c-150">В следующей процедуре показано, как получить идентификатор \***nnn \* #** выбранного слайда и использовать его в качестве целевого слайда.</span><span class="sxs-lookup"><span data-stu-id="4624c-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="4624c-151">Создайте функцию, которая получает идентификатор текущего выбранного слайда с помощью метода [Office.context.docумент. GetSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4624c-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="4624c-152">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4624c-152">The following is an example.</span></span> <span data-ttu-id="4624c-153">Обратите внимание, что вызов `getSelectedDataAsync` внедряется в функцию, возвращающую обещание.</span><span class="sxs-lookup"><span data-stu-id="4624c-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="4624c-154">Дополнительные сведения о том, почему и как это сделать, можно узнать [в статье Wrap Common-APIs в функциях, возвращающих обещаний](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="4624c-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. <span data-ttu-id="4624c-155">Вызовите новую функцию в [PowerPoint. Run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) главной функции и передайте возвращенный идентификатор (сцепленный с символом "#") в качестве значения `targetSlideId` свойства этого `InsertSlideOptions` параметра.</span><span class="sxs-lookup"><span data-stu-id="4624c-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="4624c-156">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4624c-156">The following is an example.</span></span>

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="4624c-157">Выбор слайдов для вставки</span><span class="sxs-lookup"><span data-stu-id="4624c-157">Selecting which slides to insert</span></span>

<span data-ttu-id="4624c-158">Кроме того, можно использовать параметр [инсертслидеоптионс](/javascript/api/powerpoint/powerpoint.insertslideoptions) для управления вставкой слайдов из исходной презентации.</span><span class="sxs-lookup"><span data-stu-id="4624c-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="4624c-159">Для этого необходимо назначить свойству массив идентификаторов слайдов исходной презентации `sourceSlideIds` .</span><span class="sxs-lookup"><span data-stu-id="4624c-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="4624c-160">Ниже приведен пример вставки четырех слайдов.</span><span class="sxs-lookup"><span data-stu-id="4624c-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="4624c-161">Обратите внимание, что каждая строка в массиве должна соответствовать одному или другому шаблону, используемому для `targetSlideId` Свойства.</span><span class="sxs-lookup"><span data-stu-id="4624c-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> <span data-ttu-id="4624c-162">Слайды будут вставлены в один и тот же относительный порядок, в котором они отображаются в исходной презентации независимо от того, в каком порядке они отображаются в массиве.</span><span class="sxs-lookup"><span data-stu-id="4624c-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="4624c-163">Не существует практического способа, с помощью которого пользователи могут обнаружить идентификатор или идентификатор создания слайда в исходной презентации.</span><span class="sxs-lookup"><span data-stu-id="4624c-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="4624c-164">По этой причине это свойство можно использовать только в том `sourceSlideIds` случае, если вы знаете идентификаторы источника на момент написания кода или надстройка может получить их в среде выполнения из некоторого источника данных.</span><span class="sxs-lookup"><span data-stu-id="4624c-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="4624c-165">Так как пользователи не могут запоминать идентификаторы слайдов, также необходим способ, позволяющий пользователю выбирать слайды, возможно, по названию или изображению, а затем сопоставлять каждое название или изображение с ИДЕНТИФИКАТОРом слайда.</span><span class="sxs-lookup"><span data-stu-id="4624c-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="4624c-166">Соответственно, `sourceSlideIds` свойство используется в сценариях презентации: надстройка разработана для работы с определенным набором презентаций, которые используются в качестве пулов слайдов, которые можно вставить.</span><span class="sxs-lookup"><span data-stu-id="4624c-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="4624c-167">В этом сценарии либо вы, либо клиент должны создать и поддерживать источник данных, который соответствует условию выбора (например, заголовки или изображения) с идентификаторами слайдов или идентификаторами создания слайдов, созданными из набора возможных исходных презентаций.</span><span class="sxs-lookup"><span data-stu-id="4624c-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="4624c-168">Удаление слайдов</span><span class="sxs-lookup"><span data-stu-id="4624c-168">Delete slides</span></span>

<span data-ttu-id="4624c-169">Вы можете удалить слайд, получив ссылку на объект [слайда](/javascript/api/powerpoint/powerpoint.slide) , представляющий слайд, и вызовите `Slide.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="4624c-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="4624c-170">Ниже приведен пример, в котором удаляется четвертый слайд.</span><span class="sxs-lookup"><span data-stu-id="4624c-170">The following is an example in which the 4th slide is deleted.</span></span>

```javascript
async function deleteSlide() {
  await PowerPoint.run(async function(context) {

    // The slide index is zero-based. 
    const slide = context.presentation.slides.getItemAt(3);
    slide.delete();
    await context.sync();
  });
}
```
