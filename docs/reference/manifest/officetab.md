---
title: Элемент OfficeTab в файле манифеста
description: Элемент OfficeTab определяет вкладку ленты, в которой отображается команда надстройки.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 1d1810f3d3a206f72bf9544814a3fdaaa556476e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720493"
---
# <a name="officetab-element"></a><span data-ttu-id="e6a8b-103">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e6a8b-103">OfficeTab element</span></span>

<span data-ttu-id="e6a8b-104">Определяет вкладку ленты, на которой отображается команда надстройки.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="e6a8b-105">Это может быть вкладка по умолчанию (" **домашний**", " **сообщение**" или " **собрание**") или настраиваемая вкладка, определенная надстройкой.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="e6a8b-106">Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e6a8b-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e6a8b-107">Child elements</span></span>

|  <span data-ttu-id="e6a8b-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="e6a8b-108">Element</span></span> |  <span data-ttu-id="e6a8b-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e6a8b-109">Required</span></span>  |  <span data-ttu-id="e6a8b-110">Описание</span><span class="sxs-lookup"><span data-stu-id="e6a8b-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e6a8b-111">Группа</span><span class="sxs-lookup"><span data-stu-id="e6a8b-111">Group</span></span>      | <span data-ttu-id="e6a8b-112">Да</span><span class="sxs-lookup"><span data-stu-id="e6a8b-112">Yes</span></span> |  <span data-ttu-id="e6a8b-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="e6a8b-115">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-115">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="e6a8b-116">Значения, **выделенные полужирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word в Интернете).</span><span class="sxs-lookup"><span data-stu-id="e6a8b-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="e6a8b-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="e6a8b-117">Outlook</span></span>

- <span data-ttu-id="e6a8b-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="e6a8b-119">Word</span><span class="sxs-lookup"><span data-stu-id="e6a8b-119">Word</span></span>

- <span data-ttu-id="e6a8b-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-120">**TabHome**</span></span>
- <span data-ttu-id="e6a8b-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-121">**TabInsert**</span></span>
- <span data-ttu-id="e6a8b-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="e6a8b-122">TabWordDesign</span></span>
- <span data-ttu-id="e6a8b-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="e6a8b-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="e6a8b-124">TabReferences</span></span>
- <span data-ttu-id="e6a8b-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="e6a8b-125">TabMailings</span></span>
- <span data-ttu-id="e6a8b-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="e6a8b-126">TabReviewWord</span></span>
- <span data-ttu-id="e6a8b-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-127">**TabView**</span></span>
- <span data-ttu-id="e6a8b-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e6a8b-128">TabDeveloper</span></span>
- <span data-ttu-id="e6a8b-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e6a8b-129">TabAddIns</span></span>
- <span data-ttu-id="e6a8b-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="e6a8b-130">TabBlogPost</span></span>
- <span data-ttu-id="e6a8b-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="e6a8b-131">TabBlogInsert</span></span>
- <span data-ttu-id="e6a8b-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="e6a8b-132">TabPrintPreview</span></span>
- <span data-ttu-id="e6a8b-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="e6a8b-133">TabOutlining</span></span>
- <span data-ttu-id="e6a8b-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="e6a8b-134">TabConflicts</span></span>
- <span data-ttu-id="e6a8b-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="e6a8b-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="e6a8b-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="e6a8b-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="e6a8b-137">Excel</span><span class="sxs-lookup"><span data-stu-id="e6a8b-137">Excel</span></span>

- <span data-ttu-id="e6a8b-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-138">**TabHome**</span></span>
- <span data-ttu-id="e6a8b-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-139">**TabInsert**</span></span>
- <span data-ttu-id="e6a8b-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="e6a8b-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="e6a8b-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="e6a8b-141">TabFormulas</span></span>
- <span data-ttu-id="e6a8b-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-142">**TabData**</span></span>
- <span data-ttu-id="e6a8b-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-143">**TabReview**</span></span>
- <span data-ttu-id="e6a8b-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-144">**TabView**</span></span>
- <span data-ttu-id="e6a8b-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e6a8b-145">TabDeveloper</span></span>
- <span data-ttu-id="e6a8b-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e6a8b-146">TabAddIns</span></span>
- <span data-ttu-id="e6a8b-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="e6a8b-147">TabPrintPreview</span></span>
- <span data-ttu-id="e6a8b-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="e6a8b-148">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="e6a8b-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e6a8b-149">PowerPoint</span></span>

- <span data-ttu-id="e6a8b-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-150">**TabHome**</span></span>
- <span data-ttu-id="e6a8b-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-151">**TabInsert**</span></span>
- <span data-ttu-id="e6a8b-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-152">**TabDesign**</span></span>
- <span data-ttu-id="e6a8b-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-153">**TabTransitions**</span></span>
- <span data-ttu-id="e6a8b-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-154">**TabAnimations**</span></span>
- <span data-ttu-id="e6a8b-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="e6a8b-155">TabSlideShow</span></span>
- <span data-ttu-id="e6a8b-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="e6a8b-156">TabReview</span></span>
- <span data-ttu-id="e6a8b-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-157">**TabView**</span></span>
- <span data-ttu-id="e6a8b-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e6a8b-158">TabDeveloper</span></span>
- <span data-ttu-id="e6a8b-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e6a8b-159">TabAddIns</span></span>
- <span data-ttu-id="e6a8b-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="e6a8b-160">TabPrintPreview</span></span>
- <span data-ttu-id="e6a8b-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="e6a8b-161">TabMerge</span></span>
- <span data-ttu-id="e6a8b-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="e6a8b-162">TabGrayscale</span></span>
- <span data-ttu-id="e6a8b-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="e6a8b-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="e6a8b-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="e6a8b-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="e6a8b-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="e6a8b-165">TabSlideMaster</span></span>
- <span data-ttu-id="e6a8b-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="e6a8b-166">TabHandoutMaster</span></span>
- <span data-ttu-id="e6a8b-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="e6a8b-167">TabNotesMaster</span></span>
- <span data-ttu-id="e6a8b-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="e6a8b-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="e6a8b-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="e6a8b-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="e6a8b-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="e6a8b-170">OneNote</span></span>

- <span data-ttu-id="e6a8b-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-171">**TabHome**</span></span>
- <span data-ttu-id="e6a8b-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-172">**TabInsert**</span></span>
- <span data-ttu-id="e6a8b-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e6a8b-173">**TabView**</span></span>
- <span data-ttu-id="e6a8b-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e6a8b-174">TabDeveloper</span></span>
- <span data-ttu-id="e6a8b-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e6a8b-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="e6a8b-176">Group</span><span class="sxs-lookup"><span data-stu-id="e6a8b-176">Group</span></span>

<span data-ttu-id="e6a8b-177">Группа точек расширения пользовательского интерфейса на вкладке. У группы может быть до шести элементов управления.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="e6a8b-178">Атрибут **ID** является обязательным, а каждый **идентификатор** должен быть уникальным в пределах манифеста.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="e6a8b-179">**Идентификатор** — это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="e6a8b-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="e6a8b-180">Просмотрите [элемент Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="e6a8b-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="e6a8b-181">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="e6a8b-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
