---
title: Использование пользовательских тегов на презентациях, слайдах и фигурах в PowerPoint
description: Узнайте, как использовать теги для настраиваемой метаданных о презентациях, слайдах и фигурах.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fbb13e67da1f7962fc2c0b8d45689f259b015014
ms.sourcegitcommit: 58d394fa49308ecf93cd53f7d3fb6e316ff56209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/16/2021
ms.locfileid: "51876862"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a><span data-ttu-id="84120-103">Используйте настраиваемые теги для презентаций, слайдов и фигур в PowerPoint</span><span class="sxs-lookup"><span data-stu-id="84120-103">Use custom tags for presentations, slides, and shapes in PowerPoint</span></span>

<span data-ttu-id="84120-104">Надстройка может прикреплять настраиваемые метаданные в виде пар значений ключей, называемых "тегами", к презентациям, определенным слайдам и определенным фигурам на слайде.</span><span class="sxs-lookup"><span data-stu-id="84120-104">An add-in can attach custom metadata, in the form of key-value pairs, called "tags", to presentations, specific slides, and specific shapes on a slide.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="84120-105">API для тегов находятся в предварительном просмотре.</span><span class="sxs-lookup"><span data-stu-id="84120-105">The APIs for tags are in preview.</span></span> <span data-ttu-id="84120-106">Поэкспериментируйте с ними в среде разработки или тестирования, но не добавляйте их в производственную надстройка.</span><span class="sxs-lookup"><span data-stu-id="84120-106">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>

<span data-ttu-id="84120-107">Существует два основных сценария использования тегов:</span><span class="sxs-lookup"><span data-stu-id="84120-107">There are two main scenarios for using tags:</span></span>

- <span data-ttu-id="84120-108">При применении к слайду или фигуре тег позволяет классифицовать объект для пакетной обработки.</span><span class="sxs-lookup"><span data-stu-id="84120-108">When applied to a slide or a shape, a tag enables the object to be categorized for batch processing.</span></span> <span data-ttu-id="84120-109">Например, предположим, что в презентации есть слайды, которые следует включить в презентации восточного региона, но не западного региона.</span><span class="sxs-lookup"><span data-stu-id="84120-109">For example, suppose a presentation has some slides that should be included in presentations to the East region but not the West region.</span></span> <span data-ttu-id="84120-110">Кроме того, существуют альтернативные слайды, которые должны показываться только на Западе.</span><span class="sxs-lookup"><span data-stu-id="84120-110">Similarly, there are alternative slides that should be shown only to the West.</span></span> <span data-ttu-id="84120-111">Надстройка может создать тег с ключом и значением и применить его к слайдам, которые следует `REGION` `East` использовать только на Востоке.</span><span class="sxs-lookup"><span data-stu-id="84120-111">Your add-in can create a tag with the key `REGION` and the value `East` and apply it to the slides that should only be used in the East.</span></span> <span data-ttu-id="84120-112">Значение тега заказано для слайдов, которые должны показываться только в `West` западном регионе.</span><span class="sxs-lookup"><span data-stu-id="84120-112">The tag's value is set to `West` for the slides that should only be shown to the West region.</span></span> <span data-ttu-id="84120-113">Перед презентацией на Востоке кнопка в коде надстройки выполняет циклы через все слайды, проверяя значение `REGION` тега.</span><span class="sxs-lookup"><span data-stu-id="84120-113">Just before a presentation to the East, a button in the add-in runs code that loops through all the slides checking the value of the `REGION` tag.</span></span> <span data-ttu-id="84120-114">Слайды, в которых `West` область удалена.</span><span class="sxs-lookup"><span data-stu-id="84120-114">Slides where the region is `West` are deleted.</span></span> <span data-ttu-id="84120-115">Затем пользователь закрывает надстройку и запускает слайд-шоу.</span><span class="sxs-lookup"><span data-stu-id="84120-115">The user then closes the add-in and starts the slide show.</span></span>
- <span data-ttu-id="84120-116">При применении к презентации тег фактически является настраиваемой свойством в документе презентации (аналогично [CustomProperty](/javascript/api/word/word.customproperty) в Word).</span><span class="sxs-lookup"><span data-stu-id="84120-116">When applied to a presentation, a tag is effectively a custom property in the presentation document (similar to a [CustomProperty](/javascript/api/word/word.customproperty) in Word).</span></span>

## <a name="tag-slides-and-shapes"></a><span data-ttu-id="84120-117">Слайды и фигуры тегов</span><span class="sxs-lookup"><span data-stu-id="84120-117">Tag slides and shapes</span></span>

<span data-ttu-id="84120-118">Тег — это пара значений ключа, где значение всегда типа и представлено `string` объектом [Tag.](/javascript/api/powerpoint/powerpoint.tag)</span><span class="sxs-lookup"><span data-stu-id="84120-118">A tag is a key-value pair, where the value is always of type `string` and is represented by a [Tag](/javascript/api/powerpoint/powerpoint.tag) object.</span></span> <span data-ttu-id="84120-119">Каждый тип родительского объекта, например [объект Presentation,](/javascript/api/powerpoint/powerpoint.presentation) [Slide](/javascript/api/powerpoint/powerpoint.slide)или [Shape,](/javascript/api/powerpoint/powerpoint.shape) имеет свойство `tags` типа [TagsCollection.](/javascript/api/powerpoint/powerpoint.tagcollection)</span><span class="sxs-lookup"><span data-stu-id="84120-119">Each type of parent object, such as a [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide), or [Shape](/javascript/api/powerpoint/powerpoint.shape) object, has a `tags` property of type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).</span></span>

### <a name="add-update-and-delete-tags"></a><span data-ttu-id="84120-120">Добавление, обновление и удаление тегов</span><span class="sxs-lookup"><span data-stu-id="84120-120">Add, update, and delete tags</span></span>

<span data-ttu-id="84120-121">Чтобы добавить тег к объекту, вызовите [метод TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) свойства родительского `tags` объекта.</span><span class="sxs-lookup"><span data-stu-id="84120-121">To add a tag to an object, call the [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) method of the parent object's `tags` property.</span></span> <span data-ttu-id="84120-122">Следующий код добавляет два тега к первому слайду презентации.</span><span class="sxs-lookup"><span data-stu-id="84120-122">The following code adds two tags to the first slide of a presentation.</span></span> <span data-ttu-id="84120-123">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="84120-123">About this code, note:</span></span>

- <span data-ttu-id="84120-124">Первым параметром метода `add` является ключ в паре значение ключа.</span><span class="sxs-lookup"><span data-stu-id="84120-124">The first parameter of the `add` method is the key in the key-value pair.</span></span> 
- <span data-ttu-id="84120-125">Второй параметр — это значение.</span><span class="sxs-lookup"><span data-stu-id="84120-125">The second parameter is the value.</span></span>
- <span data-ttu-id="84120-126">Ключ находится в верхних буквах.</span><span class="sxs-lookup"><span data-stu-id="84120-126">The key is in uppercase letters.</span></span> <span data-ttu-id="84120-127">Это не является строго обязательным для метода; однако ключ всегда хранится в PowerPoint в качестве верхнего шкафа, и некоторые методы, связанные с тегами, требуют, чтобы ключ был выражен в верхнем шкафу, поэтому мы рекомендуем в качестве наилучшей практики использовать верхний шкаф в коде для `add` ключа тега. </span><span class="sxs-lookup"><span data-stu-id="84120-127">This isn't strictly mandatory for the `add` method; however, the key is always stored by PowerPoint as uppercase, and *some tag-related methods do require that the key be expressed in uppercase*, so we recommend as a best practice that you always use uppercase in your code for a tag key.</span></span>

```javascript
async function addMultipleSlideTags() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}
```

<span data-ttu-id="84120-128">Метод `add` также используется для обновления тега.</span><span class="sxs-lookup"><span data-stu-id="84120-128">The `add` method is also used to update a tag.</span></span> <span data-ttu-id="84120-129">Следующий код изменяет значение `PLANET` тега.</span><span class="sxs-lookup"><span data-stu-id="84120-129">The following code changes the value of the `PLANET` tag.</span></span>

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

<span data-ttu-id="84120-130">Чтобы удалить тег, позвоните методу на родительском объекте и передайте ключ тега `delete` `TagsCollection` в качестве параметра.</span><span class="sxs-lookup"><span data-stu-id="84120-130">To delete a tag, call the `delete` method on it's parent `TagsCollection` object and pass the key of the tag as the parameter.</span></span> <span data-ttu-id="84120-131">Например, см. [в примере Set custom metadata on the presentation.](#set-custom-metadata-on-the-presentation)</span><span class="sxs-lookup"><span data-stu-id="84120-131">For an example, see [Set custom metadata on the presentation](#set-custom-metadata-on-the-presentation).</span></span>

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a><span data-ttu-id="84120-132">Использование тегов для выборочной обработки слайдов и фигур</span><span class="sxs-lookup"><span data-stu-id="84120-132">Use tags to selectively process slides and shapes</span></span>

<span data-ttu-id="84120-133">Рассмотрим следующий сценарий: Contoso Consulting имеет презентацию, которая будет показываться всем новым клиентам.</span><span class="sxs-lookup"><span data-stu-id="84120-133">Consider the following scenario: Contoso Consulting has a presentation they show to all new customers.</span></span> <span data-ttu-id="84120-134">Но некоторые слайды должны показываться только тем клиентам, которые заплатили за состояние "премиум".</span><span class="sxs-lookup"><span data-stu-id="84120-134">But some slides should only be shown to customers that have paid for "premium" status.</span></span> <span data-ttu-id="84120-135">Перед показом презентации для клиентов, не взмываюых к премиум-классам, они делают ее копию и удаляют слайды, которые должны видеть только клиенты премиум-класса.</span><span class="sxs-lookup"><span data-stu-id="84120-135">Before showing the presentation to non-premium customers, they make a copy of it and delete the slides that only premium customers should see.</span></span> <span data-ttu-id="84120-136">Надстройка позволяет Contoso теги, какие слайды для премиум-клиентов и удалить эти слайды при необходимости.</span><span class="sxs-lookup"><span data-stu-id="84120-136">An add-in enables Contoso to tag which slides are for premium customers and to delete these slides when needed.</span></span> <span data-ttu-id="84120-137">В следующем списке описаны основные этапы кодирования для создания этой функции.</span><span class="sxs-lookup"><span data-stu-id="84120-137">The following list outlines the major coding steps to create this functionality.</span></span>

1. <span data-ttu-id="84120-138">Создайте метод, который помечет выбранный в настоящее время слайд как предназначенный для `Premium` клиентов.</span><span class="sxs-lookup"><span data-stu-id="84120-138">Create a method that tags the currently selected slide as intended for `Premium` customers.</span></span> <span data-ttu-id="84120-139">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="84120-139">About this code, note:</span></span>

    - <span data-ttu-id="84120-140">Функция `getSelectedSlideIndex` определяется на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="84120-140">The `getSelectedSlideIndex` function is defined in the next step.</span></span> <span data-ttu-id="84120-141">Он возвращает индекс на основе 1 выбранного слайда.</span><span class="sxs-lookup"><span data-stu-id="84120-141">It returns the 1-based index of the currently selected slide.</span></span>
    - <span data-ttu-id="84120-142">Значение, возвращаемого функцией, должно быть отсоединимо, так как метод `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) основан на 0.</span><span class="sxs-lookup"><span data-stu-id="84120-142">The value returned by the `getSelectedSlideIndex` function has to be decremented because the [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) method is 0-based.</span></span>

    ```javascript
    async function addTagToSelectedSlide() {
      await PowerPoint.run(async function(context) {
        let selectedSlideIndex = await getSelectedSlideIndex();
        selectedSlideIndex = selectedSlideIndex - 1;
        const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
        slide.tags.add("CUSTOMER_TYPE", "Premium");
    
        await context.sync();
      });
    }
    ```

2. <span data-ttu-id="84120-143">Следующий код создает метод получения индекса выбранного слайда.</span><span class="sxs-lookup"><span data-stu-id="84120-143">The following code creates a method to get the index of the selected slide.</span></span> <span data-ttu-id="84120-144">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="84120-144">About this code, note:</span></span>

    - <span data-ttu-id="84120-145">Он использует метод [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="84120-145">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="84120-146">Вызов встроен в функцию возврата `getSelectedDataAsync` обещаний.</span><span class="sxs-lookup"><span data-stu-id="84120-146">The call to `getSelectedDataAsync` is embedded in a promise-returning function.</span></span> <span data-ttu-id="84120-147">Дополнительные сведения о том, почему и как это сделать, см. в этой ссылке [Wrap Common API в функциях возврата обещаний.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="84120-147">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="84120-148">`getSelectedDataAsync` возвращает массив, так как можно выбрать несколько слайдов.</span><span class="sxs-lookup"><span data-stu-id="84120-148">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="84120-149">В этом сценарии пользователь выбрал только один, поэтому код получает первый (0-й) слайд, который является единственным выбранным.</span><span class="sxs-lookup"><span data-stu-id="84120-149">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="84120-150">Значение слайда — это 1-основанное значение, что пользователь видит рядом со слайдом в области эскизов пользовательского интерфейса `index` PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="84120-150">The `index` value of the slide is the 1-based value the user sees beside the slide in the PowerPoint UI thumbnails pane.</span></span>

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

3. <span data-ttu-id="84120-151">В следующем коде создается метод удаления слайдов, помеченных для премиум-клиентов.</span><span class="sxs-lookup"><span data-stu-id="84120-151">The following code creates a method to delete slides that are tagged for premium customers.</span></span> <span data-ttu-id="84120-152">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="84120-152">About this code, note:</span></span>

    - <span data-ttu-id="84120-153">Так как свойства и свойства тегов будут читаться после загрузки, они должны `key` `value` быть `context.sync` загружены в первую очередь.</span><span class="sxs-lookup"><span data-stu-id="84120-153">Because the `key` and `value` properties of the tags are going to be read after the `context.sync`, they must be loaded first.</span></span>

    ```javascript
    async function deleteSlidesByAudience() {
      await PowerPoint.run(async function(context) {
        const slides = context.presentation.slides;
        slides.load("tags/key, tags/value");
    
        await context.sync();
    
        for (let i = 0; i < slides.items.length; i++) {
          let currentSlide = slides.items[i];
          for (let j = 0; j < currentSlide.tags.items.length; j++) {
            let currentTag = currentSlide.tags.items[j];
            if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
              currentSlide.delete();
            }
          }
        }
    
        await context.sync();
      });
    }
    ```

## <a name="set-custom-metadata-on-the-presentation"></a><span data-ttu-id="84120-154">Настройка настраиваемой метаданных на презентации</span><span class="sxs-lookup"><span data-stu-id="84120-154">Set custom metadata on the presentation</span></span>

<span data-ttu-id="84120-155">Надстройки также могут применять теги к презентации в целом.</span><span class="sxs-lookup"><span data-stu-id="84120-155">Add-ins can also apply tags to the presentation as a whole.</span></span> <span data-ttu-id="84120-156">Это позволяет использовать теги для метаданных на уровне документов, аналогичные использованию класса [CustomProperty](/javascript/api/word/word.customproperty)в Word.</span><span class="sxs-lookup"><span data-stu-id="84120-156">This enables you to use tags for document-level metadata similar to how the [CustomProperty](/javascript/api/word/word.customproperty)class is used in Word.</span></span> <span data-ttu-id="84120-157">Но в отличие от класса Word, значение `CustomProperty` тега PowerPoint может быть только типа `string` .</span><span class="sxs-lookup"><span data-stu-id="84120-157">But unlike the Word `CustomProperty` class, the value of a PowerPoint tag can only be of type `string`.</span></span>

<span data-ttu-id="84120-158">Следующий код — пример добавления тега в презентацию.</span><span class="sxs-lookup"><span data-stu-id="84120-158">The following code is an example of adding a tag to a presentation.</span></span> 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

<span data-ttu-id="84120-159">Следующий код — пример удаления тега из презентации.</span><span class="sxs-lookup"><span data-stu-id="84120-159">The following code is an example of deleting a tag from a presentation.</span></span> <span data-ttu-id="84120-160">Обратите внимание, что ключ тега передается `delete` методу родительского `TagsCollection` объекта.</span><span class="sxs-lookup"><span data-stu-id="84120-160">Note that the key of the tag is passed to the `delete` method of the parent `TagsCollection` object.</span></span>

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
