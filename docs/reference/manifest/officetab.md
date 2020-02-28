---
title: Элемент OfficeTab в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b8458233ba93e98fe0bd8d51f5734b1fece65864
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324836"
---
# <a name="officetab-element"></a><span data-ttu-id="c4788-102">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="c4788-102">OfficeTab element</span></span>

<span data-ttu-id="c4788-103">Определяет вкладку ленты, на которой отображается команда надстройки.</span><span class="sxs-lookup"><span data-stu-id="c4788-103">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="c4788-104">Это может быть вкладка по умолчанию (" **домашний**", " **сообщение**" или " **собрание**") или настраиваемая вкладка, определенная надстройкой.</span><span class="sxs-lookup"><span data-stu-id="c4788-104">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="c4788-105">Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="c4788-105">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c4788-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c4788-106">Child elements</span></span>

|  <span data-ttu-id="c4788-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="c4788-107">Element</span></span> |  <span data-ttu-id="c4788-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c4788-108">Required</span></span>  |  <span data-ttu-id="c4788-109">Описание</span><span class="sxs-lookup"><span data-stu-id="c4788-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c4788-110">Группа</span><span class="sxs-lookup"><span data-stu-id="c4788-110">Group</span></span>      | <span data-ttu-id="c4788-111">Да</span><span class="sxs-lookup"><span data-stu-id="c4788-111">Yes</span></span> |  <span data-ttu-id="c4788-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="c4788-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="c4788-114">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="c4788-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="c4788-115">Значения, **выделенные полужирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word в Интернете).</span><span class="sxs-lookup"><span data-stu-id="c4788-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="c4788-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="c4788-116">Outlook</span></span>

- <span data-ttu-id="c4788-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="c4788-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="c4788-118">Word</span><span class="sxs-lookup"><span data-stu-id="c4788-118">Word</span></span>

- <span data-ttu-id="c4788-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c4788-119">**TabHome**</span></span>
- <span data-ttu-id="c4788-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c4788-120">**TabInsert**</span></span>
- <span data-ttu-id="c4788-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="c4788-121">TabWordDesign</span></span>
- <span data-ttu-id="c4788-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="c4788-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="c4788-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="c4788-123">TabReferences</span></span>
- <span data-ttu-id="c4788-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="c4788-124">TabMailings</span></span>
- <span data-ttu-id="c4788-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="c4788-125">TabReviewWord</span></span>
- <span data-ttu-id="c4788-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c4788-126">**TabView**</span></span>
- <span data-ttu-id="c4788-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c4788-127">TabDeveloper</span></span>
- <span data-ttu-id="c4788-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c4788-128">TabAddIns</span></span>
- <span data-ttu-id="c4788-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="c4788-129">TabBlogPost</span></span>
- <span data-ttu-id="c4788-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="c4788-130">TabBlogInsert</span></span>
- <span data-ttu-id="c4788-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="c4788-131">TabPrintPreview</span></span>
- <span data-ttu-id="c4788-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="c4788-132">TabOutlining</span></span>
- <span data-ttu-id="c4788-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="c4788-133">TabConflicts</span></span>
- <span data-ttu-id="c4788-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="c4788-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="c4788-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="c4788-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="c4788-136">Excel</span><span class="sxs-lookup"><span data-stu-id="c4788-136">Excel</span></span>

- <span data-ttu-id="c4788-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c4788-137">**TabHome**</span></span>
- <span data-ttu-id="c4788-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c4788-138">**TabInsert**</span></span>
- <span data-ttu-id="c4788-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="c4788-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="c4788-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="c4788-140">TabFormulas</span></span>
- <span data-ttu-id="c4788-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="c4788-141">**TabData**</span></span>
- <span data-ttu-id="c4788-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="c4788-142">**TabReview**</span></span>
- <span data-ttu-id="c4788-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c4788-143">**TabView**</span></span>
- <span data-ttu-id="c4788-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c4788-144">TabDeveloper</span></span>
- <span data-ttu-id="c4788-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c4788-145">TabAddIns</span></span>
- <span data-ttu-id="c4788-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="c4788-146">TabPrintPreview</span></span>
- <span data-ttu-id="c4788-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="c4788-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="c4788-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c4788-148">PowerPoint</span></span>

- <span data-ttu-id="c4788-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c4788-149">**TabHome**</span></span>
- <span data-ttu-id="c4788-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c4788-150">**TabInsert**</span></span>
- <span data-ttu-id="c4788-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="c4788-151">**TabDesign**</span></span>
- <span data-ttu-id="c4788-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="c4788-152">**TabTransitions**</span></span>
- <span data-ttu-id="c4788-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="c4788-153">**TabAnimations**</span></span>
- <span data-ttu-id="c4788-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="c4788-154">TabSlideShow</span></span>
- <span data-ttu-id="c4788-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="c4788-155">TabReview</span></span>
- <span data-ttu-id="c4788-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c4788-156">**TabView**</span></span>
- <span data-ttu-id="c4788-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c4788-157">TabDeveloper</span></span>
- <span data-ttu-id="c4788-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c4788-158">TabAddIns</span></span>
- <span data-ttu-id="c4788-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="c4788-159">TabPrintPreview</span></span>
- <span data-ttu-id="c4788-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="c4788-160">TabMerge</span></span>
- <span data-ttu-id="c4788-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="c4788-161">TabGrayscale</span></span>
- <span data-ttu-id="c4788-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="c4788-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="c4788-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="c4788-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="c4788-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="c4788-164">TabSlideMaster</span></span>
- <span data-ttu-id="c4788-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="c4788-165">TabHandoutMaster</span></span>
- <span data-ttu-id="c4788-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="c4788-166">TabNotesMaster</span></span>
- <span data-ttu-id="c4788-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="c4788-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="c4788-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="c4788-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="c4788-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="c4788-169">OneNote</span></span>

- <span data-ttu-id="c4788-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="c4788-170">**TabHome**</span></span>
- <span data-ttu-id="c4788-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="c4788-171">**TabInsert**</span></span>
- <span data-ttu-id="c4788-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="c4788-172">**TabView**</span></span>
- <span data-ttu-id="c4788-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="c4788-173">TabDeveloper</span></span>
- <span data-ttu-id="c4788-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="c4788-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="c4788-175">Group</span><span class="sxs-lookup"><span data-stu-id="c4788-175">Group</span></span>

<span data-ttu-id="c4788-176">Группа точек расширения пользовательского интерфейса на вкладке. У группы может быть до шести элементов управления.</span><span class="sxs-lookup"><span data-stu-id="c4788-176">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="c4788-177">Атрибут **ID** является обязательным, а каждый **идентификатор** должен быть уникальным в пределах манифеста.</span><span class="sxs-lookup"><span data-stu-id="c4788-177">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="c4788-178">**Идентификатор** — это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="c4788-178">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="c4788-179">Просмотрите [элемент Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="c4788-179">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="c4788-180">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="c4788-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
