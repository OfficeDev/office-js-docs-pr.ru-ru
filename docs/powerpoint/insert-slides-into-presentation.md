---
title: Вставка и удаление слайдов в презентации PowerPoint
description: Узнайте, как вставлять слайды из одной презентации в другую и как удалять слайды.
ms.date: 01/08/2021
localization_priority: Normal
ms.openlocfilehash: a9a4b2efd1e970d9c45885f9a17046bec4de7e72
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839721"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation"></a><span data-ttu-id="b5dcd-103">Вставка и удаление слайдов в презентации PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b5dcd-103">Insert and delete slides in a PowerPoint presentation</span></span>

<span data-ttu-id="b5dcd-104">Надстройка PowerPoint может вставлять слайды из одной презентации в текущую презентацию с помощью библиотеки JavaScript для конкретного приложения PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="b5dcd-105">Можно контролировать, будут ли вставлены слайды сохранять форматирование исходных презентаций или форматирование целевой презентации.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="b5dcd-106">Вы также можете удалить слайды из презентации.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-106">You can also delete slides from the presentation.</span></span>

<span data-ttu-id="b5dcd-107">API вставки слайдов в основном используются в сценариях шаблонов презентаций: существует небольшое количество известных презентаций, которые служат в качестве пулов слайдов, которые могут быть вставлены надстройкой.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="b5dcd-108">В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки слайдов или изображения) с кодами слайдов.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="b5dcd-109">API также можно использовать в сценариях, где пользователь может вставлять слайды из любой произвольной презентации, но в этом сценарии пользователь фактически ограничивается вставкой всех слайдов из презентации источника. </span><span class="sxs-lookup"><span data-stu-id="b5dcd-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="b5dcd-110">Дополнительные сведения об этом [см.](#selecting-which-slides-to-insert) в под вопросе "Выбор слайдов для вставки".</span><span class="sxs-lookup"><span data-stu-id="b5dcd-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="b5dcd-111">Вставка слайдов из одной презентации в другую состоит из двух этапов.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="b5dcd-112">Преобразуем исходный файл презентации (PPTX) в строку в формате base64.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="b5dcd-113">Используйте этот метод, чтобы вставить один или несколько слайдов из `insertSlidesFromBase64` файла base64 в текущую презентацию.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="b5dcd-114">Преобразование исходных презентаций в base64</span><span class="sxs-lookup"><span data-stu-id="b5dcd-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="b5dcd-115">Существует множество способов преобразования файла в base64.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="b5dcd-116">Язык программирования и библиотека, которые вы используете, и будет ли преобразование на стороне сервера надстройки или на стороне клиента определяется вашим сценарием.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="b5dcd-117">Чаще всего преобразование в JavaScript происходит на стороне клиента с помощью объекта [FileReader.](https://developer.mozilla.org/docs/Web/API/FileReader)</span><span class="sxs-lookup"><span data-stu-id="b5dcd-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="b5dcd-118">В следующем примере показано, как это сделать.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="b5dcd-119">Начните с получения ссылки на исходный файл PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="b5dcd-120">В этом примере мы будем использовать тип для запроса на выбор `<input>` `file` файла пользователем.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="b5dcd-121">Добавьте следующую разметку на страницу надстройки.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="b5dcd-122">Эта разметка добавляет пользовательский интерфейс на следующий снимок экрана на страницу:</span><span class="sxs-lookup"><span data-stu-id="b5dcd-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![Screenshot showing an HTML file type input control preceded by an instructional sentence reading "Select a PowerPoint presentation from which to insert slides".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="b5dcd-125">Существует множество других способов получения файла PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="b5dcd-126">Например, если файл хранится в OneDrive или SharePoint, его можно скачать с помощью Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="b5dcd-127">Дополнительные сведения см. [в работе с файлами в Microsoft Graph](/graph/api/resources/onedrive) и access Files с помощью Microsoft [Graph.](/learn/modules/msgraph-access-file-data/)</span><span class="sxs-lookup"><span data-stu-id="b5dcd-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="b5dcd-128">Добавьте следующий код в Код JavaScript надстройки, чтобы назначить функцию событию входного `change` управления.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="b5dcd-129">(Функция `storeFileAsBase64` создается на следующем этапе.)</span><span class="sxs-lookup"><span data-stu-id="b5dcd-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="b5dcd-130">Добавьте в него указанный ниже код.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-130">Add the following code.</span></span> <span data-ttu-id="b5dcd-131">Обратите внимание на следующие вопросы об этом коде:</span><span class="sxs-lookup"><span data-stu-id="b5dcd-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="b5dcd-132">Метод `reader.readAsDataURL` преобразует файл в base64 и сохраняет его в `reader.result` свойстве.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="b5dcd-133">После завершения работы метода запускается `onload` обработок события.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="b5dcd-134">Обработник событий отключит метаданные закодированного файла и сохраняет закодированную строку `onload` в глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="b5dcd-135">Строка в кодировке base64 хранится глобально, так как она будет считываться другой функцией, которая будет создаваться на одном из последующих этапов.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="b5dcd-136">Вставка слайдов с помощью insertSlidesFromBase64</span><span class="sxs-lookup"><span data-stu-id="b5dcd-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="b5dcd-137">Надстройка вставляет слайды из другой презентации PowerPoint в текущую презентацию с помощью метода [Presentation.insertSlidesFromBase64.](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)</span><span class="sxs-lookup"><span data-stu-id="b5dcd-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="b5dcd-138">Ниже приводится простой пример, в котором все слайды из презентации источника вставляются в начале текущей презентации, а вставляемые слайды сохранят форматирование исходных файлов.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="b5dcd-139">Обратите внимание, что это глобальная переменная, которая содержит версию файла презентации PowerPoint в кодировке `chosenFileBase64` base64.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="b5dcd-140">Вы можете управлять некоторыми аспектами результата вставки, включая место вставки слайдов и то, получают ли они форматирование источника или цели, передав объект [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) в качестве второго параметра в `insertSlidesFromBase64` .</span><span class="sxs-lookup"><span data-stu-id="b5dcd-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="b5dcd-141">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-141">The following is an example.</span></span> <span data-ttu-id="b5dcd-142">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="b5dcd-142">About this code, note:</span></span>

- <span data-ttu-id="b5dcd-143">Свойство может иметь два возможных `formatting` значения: UseDestinationTheme и KeepSourceFormatting.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="b5dcd-144">При желании можно использовать `InsertSlideFormatting` это enum (например, `PowerPoint.InsertSlideFormatting.useDestinationTheme` )..</span><span class="sxs-lookup"><span data-stu-id="b5dcd-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="b5dcd-145">Функция вставляет слайды из презентации источника сразу после слайда, указанного `targetSlideId` свойством.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="b5dcd-146">Значение этого свойства является строкой одной из трех возможных форм: ***nnn\*#\*\*, \* *#* mmm***, или \**_nnn_ #* mmm\*\*\*, где *nnn* — это  ИД слайда (обычно 3 цифры), а ммм — это код создания слайда (обычно 9 цифр).</span><span class="sxs-lookup"><span data-stu-id="b5dcd-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="b5dcd-147">Некоторые примеры: `267#763315295` `267#` , и `#763315295` .</span><span class="sxs-lookup"><span data-stu-id="b5dcd-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="b5dcd-148">Конечно, во время написания кода вы, как правило, не знаете ИД или ид создания целевого слайда.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="b5dcd-149">Чаще всего надстройка просит пользователей выбрать целевой слайд.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="b5dcd-150">Далее покажите, как получить \***nnn\*#** ИД выбранного слайда и использовать его в качестве целевого слайда.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="b5dcd-151">Создайте функцию, которая получает ИД выбранного слайда с помощью метода [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="b5dcd-152">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-152">The following is an example.</span></span> <span data-ttu-id="b5dcd-153">Обратите внимание, что вызов `getSelectedDataAsync` внедрен в функцию возврата обещания.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="b5dcd-154">Дополнительные сведения о том, почему и как это сделать, см. в Common-APIs wrap в функциях возврата [обещаний.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="b5dcd-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="b5dcd-155">Вызовите новую функцию внутри [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) основной функции и передайте возвращаемую ид (вмещенную символом "#") в качестве значения свойства `targetSlideId` `InsertSlideOptions` параметра.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="b5dcd-156">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-156">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="b5dcd-157">Выбор слайдов для вставки</span><span class="sxs-lookup"><span data-stu-id="b5dcd-157">Selecting which slides to insert</span></span>

<span data-ttu-id="b5dcd-158">Вы также можете использовать параметр [InsertSlideOptions,](/javascript/api/powerpoint/powerpoint.insertslideoptions) чтобы контролировать, какие слайды из презентации источника вставляются.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="b5dcd-159">Для этого свойству назначается массив кодов слайдов презентации `sourceSlideIds` источника.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="b5dcd-160">Ниже приводится пример вставки четырех слайдов.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="b5dcd-161">Обратите внимание, что каждая строка в массиве должна следовать одному или одному из шаблонов, используемых для `targetSlideId` свойства.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="b5dcd-162">Слайды будут вставлены в том же относительном порядке, в котором они отображаются в презентации источника, независимо от порядка, в котором они отображаются в массиве.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="b5dcd-163">Не существует практического способа обнаружения пользователем ИД или ид создания слайда в презентации источника.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="b5dcd-164">По этой причине свойство можно использовать, только если вы знаете исходные коды во время кодирования или надстройка может получить их во время работы из некоторого `sourceSlideIds` источника данных.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="b5dcd-165">Так как пользователи не могут запоминать ид слайдов, вам также потребуется способ, позволяющий пользователю выбирать слайды, например по названию или изображению, а затем соотносить каждый заголовок или изображение с ид слайда.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="b5dcd-166">Соответственно, свойство в основном используется в сценариях шаблонов презентаций: надстройка предназначена для работы с определенным набором презентаций, которые выступать в качестве пулов слайдов, которые можно `sourceSlideIds` вставить.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="b5dcd-167">В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки или изображения) с кодами слайдов или идами создания слайдов, которые были сконструированы из набора возможных исходных презентаций.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="b5dcd-168">Удаление слайдов</span><span class="sxs-lookup"><span data-stu-id="b5dcd-168">Delete slides</span></span>

<span data-ttu-id="b5dcd-169">Вы можете удалить слайд, получив ссылку на объект [Slide,](/javascript/api/powerpoint/powerpoint.slide) который представляет слайд, и вызовите `Slide.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="b5dcd-170">Ниже приводится пример удаления 4-го слайда.</span><span class="sxs-lookup"><span data-stu-id="b5dcd-170">The following is an example in which the 4th slide is deleted.</span></span>

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
