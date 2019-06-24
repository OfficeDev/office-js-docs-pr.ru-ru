---
title: Элемент OfficeTab в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d073d712cec2fd58e957ffe8f344d7443d1e896e
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127564"
---
# <a name="officetab-element"></a><span data-ttu-id="93e59-102">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="93e59-102">OfficeTab element</span></span>

<span data-ttu-id="93e59-p101">Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка. Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="93e59-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="93e59-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="93e59-106">Child elements</span></span>

|  <span data-ttu-id="93e59-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="93e59-107">Element</span></span> |  <span data-ttu-id="93e59-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="93e59-108">Required</span></span>  |  <span data-ttu-id="93e59-109">Описание</span><span class="sxs-lookup"><span data-stu-id="93e59-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="93e59-110">Группа</span><span class="sxs-lookup"><span data-stu-id="93e59-110">Group</span></span>      | <span data-ttu-id="93e59-111">Да</span><span class="sxs-lookup"><span data-stu-id="93e59-111">Yes</span></span> |  <span data-ttu-id="93e59-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="93e59-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="93e59-114">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="93e59-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="93e59-115">Значения, **выделенные полужирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word в Интернете).</span><span class="sxs-lookup"><span data-stu-id="93e59-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="93e59-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="93e59-116">Outlook</span></span>

- <span data-ttu-id="93e59-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="93e59-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="93e59-118">Word</span><span class="sxs-lookup"><span data-stu-id="93e59-118">Word</span></span>

- <span data-ttu-id="93e59-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="93e59-119">**TabHome**</span></span>
- <span data-ttu-id="93e59-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="93e59-120">**TabInsert**</span></span>
- <span data-ttu-id="93e59-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="93e59-121">TabWordDesign</span></span>
- <span data-ttu-id="93e59-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="93e59-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="93e59-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="93e59-123">TabReferences</span></span>
- <span data-ttu-id="93e59-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="93e59-124">TabMailings</span></span>
- <span data-ttu-id="93e59-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="93e59-125">TabReviewWord</span></span>
- <span data-ttu-id="93e59-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="93e59-126">**TabView**</span></span>
- <span data-ttu-id="93e59-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="93e59-127">TabDeveloper</span></span>
- <span data-ttu-id="93e59-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="93e59-128">TabAddIns</span></span>
- <span data-ttu-id="93e59-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="93e59-129">TabBlogPost</span></span>
- <span data-ttu-id="93e59-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="93e59-130">TabBlogInsert</span></span>
- <span data-ttu-id="93e59-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="93e59-131">TabPrintPreview</span></span>
- <span data-ttu-id="93e59-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="93e59-132">TabOutlining</span></span>
- <span data-ttu-id="93e59-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="93e59-133">TabConflicts</span></span>
- <span data-ttu-id="93e59-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="93e59-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="93e59-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="93e59-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="93e59-136">Excel</span><span class="sxs-lookup"><span data-stu-id="93e59-136">Excel</span></span>

- <span data-ttu-id="93e59-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="93e59-137">**TabHome**</span></span>
- <span data-ttu-id="93e59-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="93e59-138">**TabInsert**</span></span>
- <span data-ttu-id="93e59-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="93e59-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="93e59-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="93e59-140">TabFormulas</span></span>
- <span data-ttu-id="93e59-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="93e59-141">**TabData**</span></span>
- <span data-ttu-id="93e59-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="93e59-142">**TabReview**</span></span>
- <span data-ttu-id="93e59-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="93e59-143">**TabView**</span></span>
- <span data-ttu-id="93e59-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="93e59-144">TabDeveloper</span></span>
- <span data-ttu-id="93e59-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="93e59-145">TabAddIns</span></span>
- <span data-ttu-id="93e59-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="93e59-146">TabPrintPreview</span></span>
- <span data-ttu-id="93e59-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="93e59-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="93e59-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="93e59-148">PowerPoint</span></span>

- <span data-ttu-id="93e59-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="93e59-149">**TabHome**</span></span>
- <span data-ttu-id="93e59-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="93e59-150">**TabInsert**</span></span>
- <span data-ttu-id="93e59-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="93e59-151">**TabDesign**</span></span>
- <span data-ttu-id="93e59-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="93e59-152">**TabTransitions**</span></span>
- <span data-ttu-id="93e59-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="93e59-153">**TabAnimations**</span></span>
- <span data-ttu-id="93e59-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="93e59-154">TabSlideShow</span></span>
- <span data-ttu-id="93e59-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="93e59-155">TabReview</span></span>
- <span data-ttu-id="93e59-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="93e59-156">**TabView**</span></span>
- <span data-ttu-id="93e59-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="93e59-157">TabDeveloper</span></span>
- <span data-ttu-id="93e59-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="93e59-158">TabAddIns</span></span>
- <span data-ttu-id="93e59-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="93e59-159">TabPrintPreview</span></span>
- <span data-ttu-id="93e59-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="93e59-160">TabMerge</span></span>
- <span data-ttu-id="93e59-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="93e59-161">TabGrayscale</span></span>
- <span data-ttu-id="93e59-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="93e59-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="93e59-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="93e59-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="93e59-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="93e59-164">TabSlideMaster</span></span>
- <span data-ttu-id="93e59-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="93e59-165">TabHandoutMaster</span></span>
- <span data-ttu-id="93e59-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="93e59-166">TabNotesMaster</span></span>
- <span data-ttu-id="93e59-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="93e59-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="93e59-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="93e59-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="93e59-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="93e59-169">OneNote</span></span>

- <span data-ttu-id="93e59-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="93e59-170">**TabHome**</span></span>
- <span data-ttu-id="93e59-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="93e59-171">**TabInsert**</span></span>
- <span data-ttu-id="93e59-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="93e59-172">**TabView**</span></span>
- <span data-ttu-id="93e59-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="93e59-173">TabDeveloper</span></span>
- <span data-ttu-id="93e59-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="93e59-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="93e59-175">Group</span><span class="sxs-lookup"><span data-stu-id="93e59-175">Group</span></span>

<span data-ttu-id="93e59-p104">Группа точек расширения пользовательского интерфейса на вкладке. В группе может быть до шести элементов управления. Атрибут **id** обязательный, и каждый атрибут **id** должен быть уникальным в манифесте. Атрибут **id** — это строка длиной до 125 символов. См. статью об[элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="93e59-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="93e59-180">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="93e59-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
