---
title: Добавление и удаление слайдов в PowerPoint
description: Узнайте, как добавлять и удалять слайды и указать мастер и макет новых слайдов.
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 9a8613997fc52ad6a30576b38c517a9c992f0e1b
ms.sourcegitcommit: ba4fb7087b9841d38bb46a99a63e88df49514a4d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/05/2021
ms.locfileid: "52779336"
---
# <a name="add-and-delete-slides-in-powerpoint"></a><span data-ttu-id="b4e42-103">Добавление и удаление слайдов в PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b4e42-103">Add and delete slides in PowerPoint</span></span>

<span data-ttu-id="b4e42-104">Надстройка PowerPoint добавить слайды в презентацию и дополнительно указать, какой мастер слайда и макет мастера используется для нового слайда.</span><span class="sxs-lookup"><span data-stu-id="b4e42-104">A PowerPoint add-in can add slides to the presentation and optionally specify which slide master, and which layout of the master, is used for the new slide.</span></span> <span data-ttu-id="b4e42-105">Надстройка также может удалять слайды.</span><span class="sxs-lookup"><span data-stu-id="b4e42-105">The add-in can also delete slides.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b4e42-106">API для добавления слайдов находятся в [предварительном просмотре](../reference/requirement-sets/powerpoint-preview-apis.md) и недоступны для надстройок производства. API для *удаления* слайдов был выпущен.</span><span class="sxs-lookup"><span data-stu-id="b4e42-106">The APIs for adding slides are in [preview](../reference/requirement-sets/powerpoint-preview-apis.md) and not available for production add-ins. The API for *deleting* slides has been released.</span></span>

<span data-ttu-id="b4e42-107">API для добавления слайдов в основном используются в сценариях, в которых коды мастеров слайдов и макеты в презентации известны во время кодирования или могут быть найдены в источнике данных во время запуска.</span><span class="sxs-lookup"><span data-stu-id="b4e42-107">The APIs for adding slides are primarily used in scenarios where the IDs of the slide masters and layouts in the presentation are known at coding time or can be found in a data source at runtime.</span></span> <span data-ttu-id="b4e42-108">В таком сценарии либо вы, либо клиент должны создать и сохранить источник данных, который сопоставляет критерий выбора (например, имена или изображения мастеров слайдов и макетов) с ID-кодами мастеров слайдов и макетов.</span><span class="sxs-lookup"><span data-stu-id="b4e42-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as the names or images of slide masters and layouts) with the IDs of the slide masters and layouts.</span></span> <span data-ttu-id="b4e42-109">API также можно использовать в сценариях, где пользователь может вставлять слайды с использованием мастера слайдов по умолчанию и макета по умолчанию, а также в сценариях, в которых пользователь может выбрать существующий слайд и создать новый с тем же мастером слайда и макетом (но не с одним и тем же контентом).</span><span class="sxs-lookup"><span data-stu-id="b4e42-109">The APIs can also be used in scenarios where the user can insert slides that use the default slide master and the master's default layout, and in scenarios where the user can select an existing slide and create a new one with the same slide master and layout (but not the same content).</span></span> <span data-ttu-id="b4e42-110">Дополнительные [сведения об этом](#selecting-which-slide-master-and-layout-to-use) см. в подборке мастера слайдов и макета.</span><span class="sxs-lookup"><span data-stu-id="b4e42-110">See [Selecting which slide master and layout to use](#selecting-which-slide-master-and-layout-to-use) for more information about this.</span></span>

## <a name="add-a-slide-with-slidecollectionadd-preview"></a><span data-ttu-id="b4e42-111">Добавление слайда с slideCollection.add (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="b4e42-111">Add a slide with SlideCollection.add (preview)</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

<span data-ttu-id="b4e42-112">Добавьте слайды [методом SlideCollection.add.](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)</span><span class="sxs-lookup"><span data-stu-id="b4e42-112">Add slides with the [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) method.</span></span> <span data-ttu-id="b4e42-113">Ниже приводится простой пример, в котором добавляется слайд, использующий мастер слайдов презентации по умолчанию и первый макет этого мастера.</span><span class="sxs-lookup"><span data-stu-id="b4e42-113">The following is a simple example in which a slide that uses the presentation's default slide master and the first layout of that master is added.</span></span> <span data-ttu-id="b4e42-114">Метод всегда добавляет новые слайды в конце презентации.</span><span class="sxs-lookup"><span data-stu-id="b4e42-114">The method always adds new slides to the end of the presentation.</span></span> <span data-ttu-id="b4e42-115">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b4e42-115">The following is an example:</span></span>

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="selecting-which-slide-master-and-layout-to-use"></a><span data-ttu-id="b4e42-116">Выбор мастера слайдов и макета для использования</span><span class="sxs-lookup"><span data-stu-id="b4e42-116">Selecting which slide master and layout to use</span></span>

<span data-ttu-id="b4e42-117">Используйте параметр [AddSlideOptions,](/javascript/api/powerpoint/powerpoint.addslideoptions) чтобы контролировать, какой мастер слайда используется для нового слайда и какой макет используется в мастере.</span><span class="sxs-lookup"><span data-stu-id="b4e42-117">Use the [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) parameter to control which slide master is used for the new slide and which layout within the master is used.</span></span> <span data-ttu-id="b4e42-118">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b4e42-118">The following is an example.</span></span> <span data-ttu-id="b4e42-119">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="b4e42-119">Note the following about this code:</span></span>

- <span data-ttu-id="b4e42-120">Вы можете включить либо оба свойства `AddSlideOptions` объекта.</span><span class="sxs-lookup"><span data-stu-id="b4e42-120">You can include either or both the properties of the `AddSlideOptions` object.</span></span>
- <span data-ttu-id="b4e42-121">Если используются оба свойства, указанный макет должен принадлежать указанному мастеру или ошибка будет выброшена.</span><span class="sxs-lookup"><span data-stu-id="b4e42-121">If both properties are used, then the specified layout must belong to the specified master or an error is thrown.</span></span>
- <span data-ttu-id="b4e42-122">Если свойство не присутствует (или его значение — пустая строка), используется мастер слайда по умолчанию и должен быть макет этого мастера `masterId` `layoutId` слайдов.</span><span class="sxs-lookup"><span data-stu-id="b4e42-122">If the `masterId` property isn't present (or its value is an empty string), then the default slide master is used and the `layoutId` must be a layout of that slide master.</span></span>
- <span data-ttu-id="b4e42-123">Мастер слайдов по умолчанию — это мастер слайдов, используемый последним слайдом в презентации.</span><span class="sxs-lookup"><span data-stu-id="b4e42-123">The default slide master is the slide master used by the last slide in the presentation.</span></span> <span data-ttu-id="b4e42-124">(В необычном случае, когда в настоящее время в презентации нет слайдов, мастер слайдов по умолчанию является первым мастером слайдов в презентации.)</span><span class="sxs-lookup"><span data-stu-id="b4e42-124">(In the unusual case where there are currently no slides in the presentation, then the default slide master is the first slide master in the presentation.)</span></span>
- <span data-ttu-id="b4e42-125">Если свойство не присутствует (или его значение — пустая строка), используется первый макет мастера, `layoutId` `masterId` заданный объектом.</span><span class="sxs-lookup"><span data-stu-id="b4e42-125">If the `layoutId` property isn't present (or its value is an empty string), then the first layout of the master that is specified by the `masterId` is used.</span></span>
- <span data-ttu-id="b4e42-126">Оба свойства являются строками одной из трех возможных форм: ***nnnnnnnn\*#\*\*, \* *#* mmmmmmmmm***, или \**_nnnnnnnnnn_ #* mmmmmmmmm\*\*\*, где *nnnnnn* is the master's or layout's ID (typically 10 digits) and *mmm* is the master's or layout's creation ID (typically 6 - 10 digits).</span><span class="sxs-lookup"><span data-stu-id="b4e42-126">Both properties are strings of one of three possible forms: \***nnnnnnnnnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnnnnnnnnn_#* mmmmmmmmm\*\*\*, where *nnnnnnnnnn* is the master's or layout's ID (typically 10 digits) and *mmmmmmmmm* is the master's or layout's creation ID (typically 6 - 10 digits).</span></span> <span data-ttu-id="b4e42-127">Некоторые примеры , `2147483690#2908289500` `2147483690#` и `#2908289500` .</span><span class="sxs-lookup"><span data-stu-id="b4e42-127">Some examples are `2147483690#2908289500`, `2147483690#`, and `#2908289500`.</span></span>

```javascript
async function addSlide() {
    await PowerPoint.run(async function(context) {
        context.presentation.slides.add({
            slideMasterId: "2147483690#2908289500",
            layoutId: "2147483691#2499880"
        });
    
        await context.sync();
    });
}
```

<span data-ttu-id="b4e42-128">Нет практических способов, чтобы пользователи могли обнаружить ID или создание ID мастера слайда или макета.</span><span class="sxs-lookup"><span data-stu-id="b4e42-128">There is no practical way that users can discover the ID or creation ID of a slide master or layout.</span></span> <span data-ttu-id="b4e42-129">По этой причине параметр можно использовать только в том случае, если вы знаете коды во время кодирования или ваша надстройка может обнаружить их `AddSlideOptions` во время работы.</span><span class="sxs-lookup"><span data-stu-id="b4e42-129">For this reason, you can really only use the `AddSlideOptions` parameter when either you know the IDs at coding time or your add-in can discover them at runtime.</span></span> <span data-ttu-id="b4e42-130">Так как нельзя ожидать, что пользователи будут запоминать ID, вам также потребуется способ, позволяющий пользователю выбрать слайды, возможно по имени или по изображению, а затем соотнести каждое название или изображение с ИД слайда.</span><span class="sxs-lookup"><span data-stu-id="b4e42-130">Because users can't be expected to memorize the IDs, you also need a way to enable the user to select slides, perhaps by name or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="b4e42-131">Соответственно, этот параметр используется в основном в сценариях, в которых надстройка предназначена для работы с определенным набором мастеров слайдов и макетов, имена которых `AddSlideOptions` известны.</span><span class="sxs-lookup"><span data-stu-id="b4e42-131">Accordingly, the `AddSlideOptions` parameter is primarily used in scenarios in which the add-in is designed to work with a specific set of slide masters and layouts whose IDs are known.</span></span> <span data-ttu-id="b4e42-132">В таком сценарии либо вы, либо клиент должны создать и сохранить источник данных, который сопоставляет критерий выбора (например, мастер слайдов и имена макетов или изображения) с соответствующими ID или кодами создания.</span><span class="sxs-lookup"><span data-stu-id="b4e42-132">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as slide master and layout names or images) with the corresponding IDs or creation IDs.</span></span>

#### <a name="have-the-user-choose-a-matching-slide"></a><span data-ttu-id="b4e42-133">Чтобы пользователь выбрал совпадающий слайд</span><span class="sxs-lookup"><span data-stu-id="b4e42-133">Have the user choose a matching slide</span></span>

<span data-ttu-id="b4e42-134">Если надстройка может использоваться в сценариях, в которых новый слайд должен использовать одно  и то же сочетание мастера слайда и макета, используемого существующим слайдом, то надстройка может (1) подсказыть пользователю выбрать слайд и (2) прочитать ID мастера слайда и макет.</span><span class="sxs-lookup"><span data-stu-id="b4e42-134">If your add-in can be used in scenarios where the new slide should use the same combination of slide master and layout that is used by an *existing* slide, then your add-in can (1) prompt the user to select a slide and (2) read the IDs of the slide master and layout.</span></span> <span data-ttu-id="b4e42-135">В следующих действиях покажите, как читать ID и добавлять слайд с мастером и макетом.</span><span class="sxs-lookup"><span data-stu-id="b4e42-135">The following steps show how to read the IDs and add a slide with a matching master and layout.</span></span>

1. <span data-ttu-id="b4e42-136">Создайте метод, чтобы получить индекс выбранного слайда.</span><span class="sxs-lookup"><span data-stu-id="b4e42-136">Create a method to get the index of the selected slide.</span></span> <span data-ttu-id="b4e42-137">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b4e42-137">The following is an example.</span></span> <span data-ttu-id="b4e42-138">Что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="b4e42-138">Note about this code:</span></span>

    - <span data-ttu-id="b4e42-139">Он использует метод [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b4e42-139">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="b4e42-140">Вызов встроен `getSelectedDataAsync` в функцию возврата обещаний.</span><span class="sxs-lookup"><span data-stu-id="b4e42-140">The call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="b4e42-141">Дополнительные сведения о том, почему и как это сделать, см. в этой ссылке [Wrap Common API в функциях возврата обещаний.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="b4e42-141">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="b4e42-142">`getSelectedDataAsync` возвращает массив, так как можно выбрать несколько слайдов.</span><span class="sxs-lookup"><span data-stu-id="b4e42-142">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="b4e42-143">В этом сценарии пользователь выбрал только один, поэтому код получает первый (0-й) слайд, который является единственным выбранным.</span><span class="sxs-lookup"><span data-stu-id="b4e42-143">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="b4e42-144">Значение `index` слайда — это 1-основанное значение, что пользователь видит рядом со слайдом в области эскизов.</span><span class="sxs-lookup"><span data-stu-id="b4e42-144">The `index` value of the slide is the 1-based value the user sees beside the slide in the thumbnails pane.</span></span>

    ```javascript
    function getSelectedSlideIndex() {
        return new OfficeExtension.Promise<number>(function(resolve, reject) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
                try {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(console.error(asyncResult.error.message));
                    } else {
                        resolve(asyncResult.value.slides[0].index);
                    }
                } 
                catch (error) {
                    reject(console.log(error));
                }
            });
        });
    }
    ```

2. <span data-ttu-id="b4e42-145">Вызов новой функции [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) основной функции, которая добавляет слайд.</span><span class="sxs-lookup"><span data-stu-id="b4e42-145">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function that adds the slide.</span></span> <span data-ttu-id="b4e42-146">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="b4e42-146">The following is an example:</span></span>

    ```javascript
    async function addSlideWithMatchingLayout() {
        await PowerPoint.run(async function(context) {
    
            let selectedSlideIndex = await getSelectedSlideIndex();
        
            // Decrement the index because the value returned by getSelectedSlideIndex()
            // is 1-based, but SlideCollection.getItemAt() is 0-based.
            const realSlideIndex = selectedSlideIndex - 1;
            const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster/id, layout/id");
        
            await context.sync();
        
            context.presentation.slides.add({
                slideMasterId: selectedSlide.slideMaster.id,
                layoutId: selectedSlide.layout.id
            });
        
            await context.sync();
        });
    }
    ```

## <a name="delete-slides"></a><span data-ttu-id="b4e42-147">Удаление слайдов</span><span class="sxs-lookup"><span data-stu-id="b4e42-147">Delete slides</span></span>

<span data-ttu-id="b4e42-148">Удалите слайд, получив ссылку на объект [Slide,](/javascript/api/powerpoint/powerpoint.slide) который представляет слайд, и позвоните по `Slide.delete` методу.</span><span class="sxs-lookup"><span data-stu-id="b4e42-148">Delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="b4e42-149">Ниже приводится пример удаления 4-го слайда:</span><span class="sxs-lookup"><span data-stu-id="b4e42-149">The following is an example in which the 4th slide is deleted:</span></span>

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
