---
title: Элемент OfficeTab в файле манифеста
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 1bf9f1d1e08a8147b52f93923229ef8fb8556fcf
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952273"
---
# <a name="officetab-element"></a><span data-ttu-id="2bdb2-102">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="2bdb2-102">OfficeTab element</span></span>

<span data-ttu-id="2bdb2-p101">Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка. Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="2bdb2-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="2bdb2-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="2bdb2-106">Child elements</span></span>

|  <span data-ttu-id="2bdb2-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="2bdb2-107">Element</span></span> |  <span data-ttu-id="2bdb2-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2bdb2-108">Required</span></span>  |  <span data-ttu-id="2bdb2-109">Описание</span><span class="sxs-lookup"><span data-stu-id="2bdb2-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2bdb2-110">Группа</span><span class="sxs-lookup"><span data-stu-id="2bdb2-110">Group</span></span>      | <span data-ttu-id="2bdb2-111">Да</span><span class="sxs-lookup"><span data-stu-id="2bdb2-111">Yes</span></span> |  <span data-ttu-id="2bdb2-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="2bdb2-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="2bdb2-114">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="2bdb2-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="2bdb2-115">Значения, **выделенные жирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word Online).</span><span class="sxs-lookup"><span data-stu-id="2bdb2-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="2bdb2-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="2bdb2-116">Outlook</span></span>

- <span data-ttu-id="2bdb2-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="2bdb2-118">Word</span><span class="sxs-lookup"><span data-stu-id="2bdb2-118">Word</span></span>

- <span data-ttu-id="2bdb2-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-119">**TabHome**</span></span>
- <span data-ttu-id="2bdb2-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-120">**TabInsert**</span></span>
- <span data-ttu-id="2bdb2-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="2bdb2-121">TabWordDesign</span></span>
- <span data-ttu-id="2bdb2-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="2bdb2-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="2bdb2-123">TabReferences</span></span>
- <span data-ttu-id="2bdb2-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="2bdb2-124">TabMailings</span></span>
- <span data-ttu-id="2bdb2-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="2bdb2-125">TabReviewWord</span></span>
- <span data-ttu-id="2bdb2-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-126">**TabView**</span></span>
- <span data-ttu-id="2bdb2-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2bdb2-127">TabDeveloper</span></span>
- <span data-ttu-id="2bdb2-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2bdb2-128">TabAddIns</span></span>
- <span data-ttu-id="2bdb2-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="2bdb2-129">TabBlogPost</span></span>
- <span data-ttu-id="2bdb2-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="2bdb2-130">TabBlogInsert</span></span>
- <span data-ttu-id="2bdb2-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2bdb2-131">TabPrintPreview</span></span>
- <span data-ttu-id="2bdb2-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="2bdb2-132">TabOutlining</span></span>
- <span data-ttu-id="2bdb2-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="2bdb2-133">TabConflicts</span></span>
- <span data-ttu-id="2bdb2-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2bdb2-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="2bdb2-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="2bdb2-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="2bdb2-136">Excel</span><span class="sxs-lookup"><span data-stu-id="2bdb2-136">Excel</span></span>

- <span data-ttu-id="2bdb2-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-137">**TabHome**</span></span>
- <span data-ttu-id="2bdb2-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-138">**TabInsert**</span></span>
- <span data-ttu-id="2bdb2-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="2bdb2-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="2bdb2-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="2bdb2-140">TabFormulas</span></span>
- <span data-ttu-id="2bdb2-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-141">**TabData**</span></span>
- <span data-ttu-id="2bdb2-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-142">**TabReview**</span></span>
- <span data-ttu-id="2bdb2-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-143">**TabView**</span></span>
- <span data-ttu-id="2bdb2-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2bdb2-144">TabDeveloper</span></span>
- <span data-ttu-id="2bdb2-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2bdb2-145">TabAddIns</span></span>
- <span data-ttu-id="2bdb2-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2bdb2-146">TabPrintPreview</span></span>
- <span data-ttu-id="2bdb2-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2bdb2-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="2bdb2-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2bdb2-148">PowerPoint</span></span>

- <span data-ttu-id="2bdb2-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-149">**TabHome**</span></span>
- <span data-ttu-id="2bdb2-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-150">**TabInsert**</span></span>
- <span data-ttu-id="2bdb2-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-151">**TabDesign**</span></span>
- <span data-ttu-id="2bdb2-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-152">**TabTransitions**</span></span>
- <span data-ttu-id="2bdb2-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-153">**TabAnimations**</span></span>
- <span data-ttu-id="2bdb2-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="2bdb2-154">TabSlideShow</span></span>
- <span data-ttu-id="2bdb2-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="2bdb2-155">TabReview</span></span>
- <span data-ttu-id="2bdb2-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-156">**TabView**</span></span>
- <span data-ttu-id="2bdb2-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2bdb2-157">TabDeveloper</span></span>
- <span data-ttu-id="2bdb2-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2bdb2-158">TabAddIns</span></span>
- <span data-ttu-id="2bdb2-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2bdb2-159">TabPrintPreview</span></span>
- <span data-ttu-id="2bdb2-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="2bdb2-160">TabMerge</span></span>
- <span data-ttu-id="2bdb2-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="2bdb2-161">TabGrayscale</span></span>
- <span data-ttu-id="2bdb2-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="2bdb2-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="2bdb2-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="2bdb2-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="2bdb2-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="2bdb2-164">TabSlideMaster</span></span>
- <span data-ttu-id="2bdb2-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="2bdb2-165">TabHandoutMaster</span></span>
- <span data-ttu-id="2bdb2-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="2bdb2-166">TabNotesMaster</span></span>
- <span data-ttu-id="2bdb2-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2bdb2-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="2bdb2-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="2bdb2-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="2bdb2-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="2bdb2-169">OneNote</span></span>

- <span data-ttu-id="2bdb2-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-170">**TabHome**</span></span>
- <span data-ttu-id="2bdb2-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-171">**TabInsert**</span></span>
- <span data-ttu-id="2bdb2-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2bdb2-172">**TabView**</span></span>
- <span data-ttu-id="2bdb2-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2bdb2-173">TabDeveloper</span></span>
- <span data-ttu-id="2bdb2-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2bdb2-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="2bdb2-175">Group</span><span class="sxs-lookup"><span data-stu-id="2bdb2-175">Group</span></span>

<span data-ttu-id="2bdb2-p104">Группа точек расширения пользовательского интерфейса на вкладке. В группе может быть до шести элементов управления. Атрибут **id** обязательный, и каждый атрибут **id** должен быть уникальным в манифесте. Атрибут **id** — это строка длиной до 125 символов. См. статью об[элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="2bdb2-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="2bdb2-180">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="2bdb2-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
