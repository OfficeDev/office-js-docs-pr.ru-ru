---
title: Вставка слайдов в PowerPoint презентации
description: Узнайте, как вставить слайды из одной презентации в другую.
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 9b106e8940e7b0f19678e0467d8e900ffecd9438
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348785"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a><span data-ttu-id="f9b3f-103">Вставка слайдов в PowerPoint презентации</span><span class="sxs-lookup"><span data-stu-id="f9b3f-103">Insert slides in a PowerPoint presentation</span></span>

<span data-ttu-id="f9b3f-104">Надстройка PowerPoint может вставлять слайды из одной презентации в текущую презентацию с помощью PowerPoint библиотеки JavaScript, определенной для приложений.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="f9b3f-105">Вы можете контролировать, будут ли вставлены слайды сохранять форматирование исходных презентаций или форматирование целевой презентации.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span>

<span data-ttu-id="f9b3f-106">API вставки слайдов в основном используются в сценариях шаблонов презентаций: существует небольшое количество известных презентаций, которые служат пулами слайдов, которые могут быть вставлены надстройкой.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-106">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="f9b3f-107">В таком сценарии либо вы, либо клиент должны создать и сохранить источник данных, который сопоставляет критерий выбора (например, заголовки слайдов или изображения) с кодами слайдов.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-107">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="f9b3f-108">API также можно использовать в сценариях, в которых пользователь может вставлять слайды из любой произвольной  презентации, но в этом случае пользователь фактически ограничивается вставкой всех слайдов из исходных презентаций.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-108">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="f9b3f-109">Дополнительные [сведения об этом](#selecting-which-slides-to-insert) см. в подборе слайдов, которые необходимо вставить.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-109">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="f9b3f-110">Существует два шага к вставке слайдов из одной презентации в другую.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-110">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="f9b3f-111">Преобразование файла исходных презентаций (.pptx) в строку с форматом base64.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-111">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="f9b3f-112">Используйте `insertSlidesFromBase64` метод, чтобы вставить один или несколько слайдов из файла base64 в текущую презентацию.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-112">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="f9b3f-113">Преобразование исходных презентаций в base64</span><span class="sxs-lookup"><span data-stu-id="f9b3f-113">Convert the source presentation to base64</span></span>

<span data-ttu-id="f9b3f-114">Существует множество способов преобразования файла в base64.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-114">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="f9b3f-115">Язык программирования и библиотека, которые вы используете, и преобразование на серверной стороне надстройки или клиентской стороне определяется вашим сценарием.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-115">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="f9b3f-116">Чаще всего преобразование в JavaScript будет происходить с клиентской стороны с помощью объекта [FileReader.](https://developer.mozilla.org/docs/Web/API/FileReader)</span><span class="sxs-lookup"><span data-stu-id="f9b3f-116">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="f9b3f-117">В следующем примере показана эта практика.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-117">The following example shows this practice.</span></span>

1. <span data-ttu-id="f9b3f-118">Начните с получения ссылки на исходный PowerPoint файл.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-118">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="f9b3f-119">В этом примере мы будем использовать управление типом, чтобы побудить пользователя `<input>` `file` выбрать файл.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-119">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="f9b3f-120">Добавьте следующую разметку на страницу надстройки.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-120">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="f9b3f-121">Эта разметка добавляет пользовательский интерфейс на следующий скриншот страницы.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-121">This markup adds the UI in the following screenshot to the page.</span></span>

    ![Снимок экрана, показывающий элемент управления вводом типа HTML-файлов, предшествующего инструкции по чтению предложения "Выберите презентацию PowerPoint, из которой вставить слайды".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="f9b3f-124">Существует множество других способов получения PowerPoint файла.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-124">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="f9b3f-125">Например, если файл хранится на OneDrive или SharePoint, для его Graph Microsoft.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-125">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="f9b3f-126">Дополнительные сведения см. в материалах [Working with files in Microsoft Graph](/graph/api/resources/onedrive) и Access Files with Microsoft [Graph.](/learn/modules/msgraph-access-file-data/)</span><span class="sxs-lookup"><span data-stu-id="f9b3f-126">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="f9b3f-127">Добавьте следующий код в JavaScript надстройки, чтобы назначить функцию событию управления `change` входом.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-127">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="f9b3f-128">(Вы создаете `storeFileAsBase64` функцию на следующем шаге.)</span><span class="sxs-lookup"><span data-stu-id="f9b3f-128">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="f9b3f-129">Добавьте в него указанный ниже код.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-129">Add the following code.</span></span> <span data-ttu-id="f9b3f-130">Обратите внимание на следующие аспекты этого кода.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-130">Note the following about this code.</span></span>

    - <span data-ttu-id="f9b3f-131">Метод `reader.readAsDataURL` преобразует файл в base64 и сохраняет его в `reader.result` свойстве.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-131">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="f9b3f-132">Когда метод завершается, он запускает `onload` обработник событий.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-132">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="f9b3f-133">Обработник событий отделяет метаданные от закодированного файла и сохраняет закодированную строку `onload` в глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-133">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="f9b3f-134">Строка с кодом base64 хранится глобально, так как она будет считываться другой функцией, которую вы создаете на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-134">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="f9b3f-135">Вставка слайдов со вставкамиSlidesFromBase64</span><span class="sxs-lookup"><span data-stu-id="f9b3f-135">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="f9b3f-136">Ваша надстройка вставляет слайды из другого PowerPoint в текущую презентацию с помощью метода [Presentation.insertSlidesFromBase64.](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)</span><span class="sxs-lookup"><span data-stu-id="f9b3f-136">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="f9b3f-137">Ниже приводится простой пример, в котором все слайды из презентации источника вставляются в начале текущей презентации, а вставленные слайды держат форматирование исходных файлов.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-137">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="f9b3f-138">Обратите внимание, что это глобальная переменная, которая содержит базовую версию файла PowerPoint `chosenFileBase64` презентации.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-138">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="f9b3f-139">Вы можете управлять некоторыми аспектами результата вставки, в том числе с помощью вставки слайдов и получения источника или целевого форматирования, передав объект [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) в качестве второго параметра `insertSlidesFromBase64` .</span><span class="sxs-lookup"><span data-stu-id="f9b3f-139">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="f9b3f-140">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-140">The following is an example.</span></span> <span data-ttu-id="f9b3f-141">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="f9b3f-141">About this code, note:</span></span>

- <span data-ttu-id="f9b3f-142">Существует два возможных значения `formatting` свойства: "UseDestinationTheme" и "KeepSourceFormatting".</span><span class="sxs-lookup"><span data-stu-id="f9b3f-142">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="f9b3f-143">Необязательный, вы можете использовать `InsertSlideFormatting` enum, (например, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).</span><span class="sxs-lookup"><span data-stu-id="f9b3f-143">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="f9b3f-144">Функция будет вставлять слайды из презентации источника сразу после слайда, указанного `targetSlideId` свойством.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-144">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="f9b3f-145">Значение этого свойства — строка из трех возможных форм: ***nnn\*#\*\*, \* *#* mmmmmmmmm***, или \**_nnn_ #* mmmmmmmmm\*\*\*, где nnn — это *ID* слайда (обычно 3 цифры), а *mmmmmmmmmmm* — это код создания слайда (обычно 9 цифр).</span><span class="sxs-lookup"><span data-stu-id="f9b3f-145">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="f9b3f-146">Некоторые примеры , `267#763315295` `267#` и `#763315295` .</span><span class="sxs-lookup"><span data-stu-id="f9b3f-146">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="f9b3f-147">Конечно, во время кодирования обычно не будет знать ID или код создания целевого слайда.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-147">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="f9b3f-148">Чаще всего надстройка будет просить пользователей выбрать целевой слайд.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-148">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="f9b3f-149">В следующих действиях покажите, как получить \***nnn\*#** ID выбранного в настоящее время слайда и использовать его в качестве целевого слайда.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-149">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="f9b3f-150">Создайте функцию, которая получает ID выбранного в настоящее время слайда с помощью метода [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-150">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="f9b3f-151">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-151">The following is an example.</span></span> <span data-ttu-id="f9b3f-152">Обратите внимание, что вызов `getSelectedDataAsync` встроен в функцию возврата обещаний.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-152">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="f9b3f-153">Дополнительные сведения о том, почему и как это сделать, см. в Common-APIs wrap Common-APIs в функциях возврата [обещаний.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="f9b3f-153">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="f9b3f-154">Вызов новой функции [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) главной функции и передать возвращаемую (сопутованую символу #) ID в качестве значения свойства `targetSlideId` `InsertSlideOptions` параметра.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-154">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="f9b3f-155">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-155">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="f9b3f-156">Выбор слайдов для вставки</span><span class="sxs-lookup"><span data-stu-id="f9b3f-156">Selecting which slides to insert</span></span>

<span data-ttu-id="f9b3f-157">Вы также можете использовать параметр [InsertSlideOptions,](/javascript/api/powerpoint/powerpoint.insertslideoptions) чтобы контролировать, какие слайды из презентации источника вставляются.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-157">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="f9b3f-158">Это необходимо, назначив свойству массив слайд-кодов исходных `sourceSlideIds` презентаций.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-158">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="f9b3f-159">Ниже приводится пример, в который вставляется четыре слайда.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-159">The following is an example that inserts four slides.</span></span> <span data-ttu-id="f9b3f-160">Обратите внимание, что каждая строка в массиве должна следовать тем или иным шаблонам, используемым для `targetSlideId` свойства.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-160">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="f9b3f-161">Слайды будут вставлены в том же относительном порядке, в котором они отображаются в презентации источника, независимо от порядка, в котором они отображаются в массиве.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-161">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="f9b3f-162">Нет практического способа, чтобы пользователи могли обнаружить ID или код создания слайда в презентации источника.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-162">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="f9b3f-163">По этой причине свойство можно использовать только в том случае, если вы знаете исходные коды во время кодирования или ваша надстройка может получить их в период работы из источника `sourceSlideIds` данных.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-163">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="f9b3f-164">Поскольку нельзя ожидать, что пользователи будут запоминать слайд-ИД, вам также необходим способ, позволяющий пользователю выбрать слайды, возможно, по названию или по изображению, а затем соотнести каждое название или изображение с ИД слайда.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-164">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="f9b3f-165">Соответственно, свойство используется в основном в сценариях шаблонов презентаций: надстройка предназначена для работы с определенным набором презентаций, которые служат пулами слайдов, которые можно `sourceSlideIds` вставить.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-165">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="f9b3f-166">В таком сценарии либо вы, либо клиент должны создавать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки или изображения) с кодами слайдов или кодами создания слайдов, которые были построены из набора возможных исходных презентаций.</span><span class="sxs-lookup"><span data-stu-id="f9b3f-166">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>
