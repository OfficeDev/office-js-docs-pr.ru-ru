# <a name="officetab-element"></a><span data-ttu-id="b0193-101">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b0193-101">OfficeTab element</span></span>

<span data-ttu-id="b0193-p101">Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка. Этот элемент обязателен.</span><span class="sxs-lookup"><span data-stu-id="b0193-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b0193-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="b0193-105">Child elements</span></span>

|  <span data-ttu-id="b0193-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="b0193-106">Element</span></span> |  <span data-ttu-id="b0193-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="b0193-107">Required</span></span>  |  <span data-ttu-id="b0193-108">Описание</span><span class="sxs-lookup"><span data-stu-id="b0193-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b0193-109">Group</span><span class="sxs-lookup"><span data-stu-id="b0193-109">Group</span></span>      | <span data-ttu-id="b0193-110">Да</span><span class="sxs-lookup"><span data-stu-id="b0193-110">Yes</span></span> |  <span data-ttu-id="b0193-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="b0193-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="b0193-113">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="b0193-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="b0193-114">Значения, выделенные **полужирным шрифтом**, поддерживаются классическими и веб-приложениями (например, Word 2016 для Windows и Word Online).</span><span class="sxs-lookup"><span data-stu-id="b0193-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="b0193-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="b0193-115">Outlook</span></span>

- <span data-ttu-id="b0193-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="b0193-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="b0193-117">Word</span><span class="sxs-lookup"><span data-stu-id="b0193-117">Word</span></span>

- <span data-ttu-id="b0193-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b0193-118">**TabHome**</span></span>
- <span data-ttu-id="b0193-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b0193-119">**TabInsert**</span></span>
- <span data-ttu-id="b0193-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="b0193-120">TabWordDesign</span></span>
- <span data-ttu-id="b0193-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="b0193-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="b0193-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="b0193-122">TabReferences</span></span>
- <span data-ttu-id="b0193-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="b0193-123">TabMailings</span></span>
- <span data-ttu-id="b0193-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="b0193-124">TabReviewWord</span></span>
- <span data-ttu-id="b0193-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b0193-125">**TabView**</span></span>
- <span data-ttu-id="b0193-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b0193-126">TabDeveloper</span></span>
- <span data-ttu-id="b0193-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b0193-127">TabAddIns</span></span>
- <span data-ttu-id="b0193-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="b0193-128">TabBlogPost</span></span>
- <span data-ttu-id="b0193-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="b0193-129">TabBlogInsert</span></span>
- <span data-ttu-id="b0193-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="b0193-130">TabPrintPreview</span></span>
- <span data-ttu-id="b0193-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="b0193-131">TabOutlining</span></span>
- <span data-ttu-id="b0193-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="b0193-132">TabConflicts</span></span>
- <span data-ttu-id="b0193-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="b0193-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="b0193-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="b0193-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="b0193-135">Excel</span><span class="sxs-lookup"><span data-stu-id="b0193-135">Excel</span></span>

- <span data-ttu-id="b0193-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b0193-136">**TabHome**</span></span>
- <span data-ttu-id="b0193-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b0193-137">**TabInsert**</span></span>
- <span data-ttu-id="b0193-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="b0193-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="b0193-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="b0193-139">TabFormulas</span></span>
- <span data-ttu-id="b0193-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="b0193-140">**TabData**</span></span>
- <span data-ttu-id="b0193-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="b0193-141">**TabReview**</span></span>
- <span data-ttu-id="b0193-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b0193-142">**TabView**</span></span>
- <span data-ttu-id="b0193-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b0193-143">TabDeveloper</span></span>
- <span data-ttu-id="b0193-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b0193-144">TabAddIns</span></span>
- <span data-ttu-id="b0193-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="b0193-145">TabPrintPreview</span></span>
- <span data-ttu-id="b0193-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="b0193-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="b0193-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b0193-147">PowerPoint</span></span>

- <span data-ttu-id="b0193-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b0193-148">**TabHome**</span></span>
- <span data-ttu-id="b0193-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b0193-149">**TabInsert**</span></span>
- <span data-ttu-id="b0193-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="b0193-150">**TabDesign**</span></span>
- <span data-ttu-id="b0193-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="b0193-151">**TabTransitions**</span></span>
- <span data-ttu-id="b0193-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="b0193-152">**TabAnimations**</span></span>
- <span data-ttu-id="b0193-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="b0193-153">TabSlideShow</span></span>
- <span data-ttu-id="b0193-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="b0193-154">TabReview</span></span>
- <span data-ttu-id="b0193-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b0193-155">**TabView**</span></span>
- <span data-ttu-id="b0193-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b0193-156">TabDeveloper</span></span>
- <span data-ttu-id="b0193-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b0193-157">TabAddIns</span></span>
- <span data-ttu-id="b0193-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="b0193-158">TabPrintPreview</span></span>
- <span data-ttu-id="b0193-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="b0193-159">TabMerge</span></span>
- <span data-ttu-id="b0193-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="b0193-160">TabGrayscale</span></span>
- <span data-ttu-id="b0193-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="b0193-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="b0193-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="b0193-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="b0193-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="b0193-163">TabSlideMaster</span></span>
- <span data-ttu-id="b0193-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="b0193-164">TabHandoutMaster</span></span>
- <span data-ttu-id="b0193-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="b0193-165">TabNotesMaster</span></span>
- <span data-ttu-id="b0193-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="b0193-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="b0193-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="b0193-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="b0193-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="b0193-168">OneNote</span></span>

- <span data-ttu-id="b0193-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b0193-169">**TabHome**</span></span>
- <span data-ttu-id="b0193-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b0193-170">**TabInsert**</span></span>
- <span data-ttu-id="b0193-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b0193-171">**TabView**</span></span>
- <span data-ttu-id="b0193-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b0193-172">TabDeveloper</span></span>
- <span data-ttu-id="b0193-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b0193-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="b0193-174">Group</span><span class="sxs-lookup"><span data-stu-id="b0193-174">Group</span></span>

<span data-ttu-id="b0193-p104">Группа точек расширения пользовательского интерфейса на вкладке. В группе может быть до шести элементов управления. Атрибут **id** обязательный, и каждый атрибут **id** должен быть уникальным в манифесте. Атрибут **id** — это строка длиной до 125 символов. См. статью [Элемент Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="b0193-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="b0193-179">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="b0193-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
