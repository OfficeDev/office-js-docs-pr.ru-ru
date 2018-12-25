---
title: Элемент OfficeTab в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 721064687c3c892b565a94e418815726cc0817f5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432874"
---
# <a name="officetab-element"></a><span data-ttu-id="dfcc1-102">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="dfcc1-102">OfficeTab element</span></span>

<span data-ttu-id="dfcc1-p101">Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка. Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="dfcc1-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dfcc1-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="dfcc1-106">Child elements</span></span>

|  <span data-ttu-id="dfcc1-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="dfcc1-107">Element</span></span> |  <span data-ttu-id="dfcc1-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="dfcc1-108">Required</span></span>  |  <span data-ttu-id="dfcc1-109">Описание</span><span class="sxs-lookup"><span data-stu-id="dfcc1-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="dfcc1-110">Group</span><span class="sxs-lookup"><span data-stu-id="dfcc1-110">Group</span></span>      | <span data-ttu-id="dfcc1-111">Да</span><span class="sxs-lookup"><span data-stu-id="dfcc1-111">Yes</span></span> |  <span data-ttu-id="dfcc1-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="dfcc1-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="dfcc1-114">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="dfcc1-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="dfcc1-115">Значения, выделенные **полужирным шрифтом**, поддерживаются классическими и веб-приложениями (например, Word 2016 или более поздней версии для Windows и Word Online).</span><span class="sxs-lookup"><span data-stu-id="dfcc1-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="dfcc1-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="dfcc1-116">Outlook</span></span>

- <span data-ttu-id="dfcc1-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="dfcc1-118">Word</span><span class="sxs-lookup"><span data-stu-id="dfcc1-118">Word</span></span>

- <span data-ttu-id="dfcc1-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-119">**TabHome**</span></span>
- <span data-ttu-id="dfcc1-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-120">**TabInsert**</span></span>
- <span data-ttu-id="dfcc1-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="dfcc1-121">TabWordDesign</span></span>
- <span data-ttu-id="dfcc1-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="dfcc1-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="dfcc1-123">TabReferences</span></span>
- <span data-ttu-id="dfcc1-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="dfcc1-124">TabMailings</span></span>
- <span data-ttu-id="dfcc1-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="dfcc1-125">TabReviewWord</span></span>
- <span data-ttu-id="dfcc1-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-126">**TabView**</span></span>
- <span data-ttu-id="dfcc1-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dfcc1-127">TabDeveloper</span></span>
- <span data-ttu-id="dfcc1-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dfcc1-128">TabAddIns</span></span>
- <span data-ttu-id="dfcc1-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="dfcc1-129">TabBlogPost</span></span>
- <span data-ttu-id="dfcc1-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="dfcc1-130">TabBlogInsert</span></span>
- <span data-ttu-id="dfcc1-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dfcc1-131">TabPrintPreview</span></span>
- <span data-ttu-id="dfcc1-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="dfcc1-132">TabOutlining</span></span>
- <span data-ttu-id="dfcc1-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="dfcc1-133">TabConflicts</span></span>
- <span data-ttu-id="dfcc1-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dfcc1-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="dfcc1-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="dfcc1-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="dfcc1-136">Excel</span><span class="sxs-lookup"><span data-stu-id="dfcc1-136">Excel</span></span>

- <span data-ttu-id="dfcc1-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-137">**TabHome**</span></span>
- <span data-ttu-id="dfcc1-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-138">**TabInsert**</span></span>
- <span data-ttu-id="dfcc1-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="dfcc1-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="dfcc1-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="dfcc1-140">TabFormulas</span></span>
- <span data-ttu-id="dfcc1-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-141">**TabData**</span></span>
- <span data-ttu-id="dfcc1-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-142">**TabReview**</span></span>
- <span data-ttu-id="dfcc1-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-143">**TabView**</span></span>
- <span data-ttu-id="dfcc1-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dfcc1-144">TabDeveloper</span></span>
- <span data-ttu-id="dfcc1-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dfcc1-145">TabAddIns</span></span>
- <span data-ttu-id="dfcc1-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dfcc1-146">TabPrintPreview</span></span>
- <span data-ttu-id="dfcc1-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dfcc1-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="dfcc1-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="dfcc1-148">PowerPoint</span></span>

- <span data-ttu-id="dfcc1-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-149">**TabHome**</span></span>
- <span data-ttu-id="dfcc1-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-150">**TabInsert**</span></span>
- <span data-ttu-id="dfcc1-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-151">**TabDesign**</span></span>
- <span data-ttu-id="dfcc1-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-152">**TabTransitions**</span></span>
- <span data-ttu-id="dfcc1-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-153">**TabAnimations**</span></span>
- <span data-ttu-id="dfcc1-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="dfcc1-154">TabSlideShow</span></span>
- <span data-ttu-id="dfcc1-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="dfcc1-155">TabReview</span></span>
- <span data-ttu-id="dfcc1-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-156">**TabView**</span></span>
- <span data-ttu-id="dfcc1-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dfcc1-157">TabDeveloper</span></span>
- <span data-ttu-id="dfcc1-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dfcc1-158">TabAddIns</span></span>
- <span data-ttu-id="dfcc1-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dfcc1-159">TabPrintPreview</span></span>
- <span data-ttu-id="dfcc1-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="dfcc1-160">TabMerge</span></span>
- <span data-ttu-id="dfcc1-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="dfcc1-161">TabGrayscale</span></span>
- <span data-ttu-id="dfcc1-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="dfcc1-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="dfcc1-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="dfcc1-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="dfcc1-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="dfcc1-164">TabSlideMaster</span></span>
- <span data-ttu-id="dfcc1-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="dfcc1-165">TabHandoutMaster</span></span>
- <span data-ttu-id="dfcc1-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="dfcc1-166">TabNotesMaster</span></span>
- <span data-ttu-id="dfcc1-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dfcc1-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="dfcc1-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="dfcc1-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="dfcc1-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="dfcc1-169">OneNote</span></span>

- <span data-ttu-id="dfcc1-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-170">**TabHome**</span></span>
- <span data-ttu-id="dfcc1-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-171">**TabInsert**</span></span>
- <span data-ttu-id="dfcc1-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dfcc1-172">**TabView**</span></span>
- <span data-ttu-id="dfcc1-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dfcc1-173">TabDeveloper</span></span>
- <span data-ttu-id="dfcc1-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dfcc1-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="dfcc1-175">Group</span><span class="sxs-lookup"><span data-stu-id="dfcc1-175">Group</span></span>

<span data-ttu-id="dfcc1-p104">Группа точек расширения пользовательского интерфейса на вкладке. В группе может быть до шести элементов управления. Атрибут **id** обязательный, и каждый атрибут **id** должен быть уникальным в манифесте. Атрибут **id** — это строка длиной до 125 символов. См. статью об[элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="dfcc1-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="dfcc1-180">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="dfcc1-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
