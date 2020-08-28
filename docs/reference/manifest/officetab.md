---
title: Элемент OfficeTab в файле манифеста
description: Элемент OfficeTab определяет вкладку ленты, в которой отображается команда надстройки.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 25e8044d8b3264bf9ee64c54487566bf11f0065e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292302"
---
# <a name="officetab-element"></a><span data-ttu-id="aee4b-103">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="aee4b-103">OfficeTab element</span></span>

<span data-ttu-id="aee4b-104">Определяет вкладку ленты, на которой отображается команда надстройки.</span><span class="sxs-lookup"><span data-stu-id="aee4b-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="aee4b-105">Это может быть вкладка по умолчанию (" **домашний**", " **сообщение**" или " **собрание**") или настраиваемая вкладка, определенная надстройкой.</span><span class="sxs-lookup"><span data-stu-id="aee4b-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="aee4b-106">Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="aee4b-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="aee4b-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="aee4b-107">Child elements</span></span>

|  <span data-ttu-id="aee4b-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="aee4b-108">Element</span></span> |  <span data-ttu-id="aee4b-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="aee4b-109">Required</span></span>  |  <span data-ttu-id="aee4b-110">Описание</span><span class="sxs-lookup"><span data-stu-id="aee4b-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="aee4b-111">Группа</span><span class="sxs-lookup"><span data-stu-id="aee4b-111">Group</span></span>      | <span data-ttu-id="aee4b-112">Да</span><span class="sxs-lookup"><span data-stu-id="aee4b-112">Yes</span></span> |  <span data-ttu-id="aee4b-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="aee4b-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="aee4b-115">Ниже приведены допустимые `id` значения вкладок приложения.</span><span class="sxs-lookup"><span data-stu-id="aee4b-115">The following are valid tab `id` values by application.</span></span> <span data-ttu-id="aee4b-116">Значения, **выделенные полужирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word в Интернете).</span><span class="sxs-lookup"><span data-stu-id="aee4b-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="aee4b-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="aee4b-117">Outlook</span></span>

- <span data-ttu-id="aee4b-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="aee4b-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="aee4b-119">Word</span><span class="sxs-lookup"><span data-stu-id="aee4b-119">Word</span></span>

- <span data-ttu-id="aee4b-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="aee4b-120">**TabHome**</span></span>
- <span data-ttu-id="aee4b-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="aee4b-121">**TabInsert**</span></span>
- <span data-ttu-id="aee4b-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="aee4b-122">TabWordDesign</span></span>
- <span data-ttu-id="aee4b-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="aee4b-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="aee4b-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="aee4b-124">TabReferences</span></span>
- <span data-ttu-id="aee4b-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="aee4b-125">TabMailings</span></span>
- <span data-ttu-id="aee4b-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="aee4b-126">TabReviewWord</span></span>
- <span data-ttu-id="aee4b-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="aee4b-127">**TabView**</span></span>
- <span data-ttu-id="aee4b-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="aee4b-128">TabDeveloper</span></span>
- <span data-ttu-id="aee4b-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="aee4b-129">TabAddIns</span></span>
- <span data-ttu-id="aee4b-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="aee4b-130">TabBlogPost</span></span>
- <span data-ttu-id="aee4b-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="aee4b-131">TabBlogInsert</span></span>
- <span data-ttu-id="aee4b-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="aee4b-132">TabPrintPreview</span></span>
- <span data-ttu-id="aee4b-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="aee4b-133">TabOutlining</span></span>
- <span data-ttu-id="aee4b-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="aee4b-134">TabConflicts</span></span>
- <span data-ttu-id="aee4b-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="aee4b-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="aee4b-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="aee4b-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="aee4b-137">Excel</span><span class="sxs-lookup"><span data-stu-id="aee4b-137">Excel</span></span>

- <span data-ttu-id="aee4b-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="aee4b-138">**TabHome**</span></span>
- <span data-ttu-id="aee4b-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="aee4b-139">**TabInsert**</span></span>
- <span data-ttu-id="aee4b-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="aee4b-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="aee4b-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="aee4b-141">TabFormulas</span></span>
- <span data-ttu-id="aee4b-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="aee4b-142">**TabData**</span></span>
- <span data-ttu-id="aee4b-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="aee4b-143">**TabReview**</span></span>
- <span data-ttu-id="aee4b-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="aee4b-144">**TabView**</span></span>
- <span data-ttu-id="aee4b-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="aee4b-145">TabDeveloper</span></span>
- <span data-ttu-id="aee4b-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="aee4b-146">TabAddIns</span></span>
- <span data-ttu-id="aee4b-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="aee4b-147">TabPrintPreview</span></span>
- <span data-ttu-id="aee4b-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="aee4b-148">TabBackgroundRemoval</span></span>

### <a name="powerpoint"></a><span data-ttu-id="aee4b-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="aee4b-149">PowerPoint</span></span>

- <span data-ttu-id="aee4b-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="aee4b-150">**TabHome**</span></span>
- <span data-ttu-id="aee4b-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="aee4b-151">**TabInsert**</span></span>
- <span data-ttu-id="aee4b-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="aee4b-152">**TabDesign**</span></span>
- <span data-ttu-id="aee4b-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="aee4b-153">**TabTransitions**</span></span>
- <span data-ttu-id="aee4b-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="aee4b-154">**TabAnimations**</span></span>
- <span data-ttu-id="aee4b-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="aee4b-155">TabSlideShow</span></span>
- <span data-ttu-id="aee4b-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="aee4b-156">TabReview</span></span>
- <span data-ttu-id="aee4b-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="aee4b-157">**TabView**</span></span>
- <span data-ttu-id="aee4b-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="aee4b-158">TabDeveloper</span></span>
- <span data-ttu-id="aee4b-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="aee4b-159">TabAddIns</span></span>
- <span data-ttu-id="aee4b-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="aee4b-160">TabPrintPreview</span></span>
- <span data-ttu-id="aee4b-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="aee4b-161">TabMerge</span></span>
- <span data-ttu-id="aee4b-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="aee4b-162">TabGrayscale</span></span>
- <span data-ttu-id="aee4b-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="aee4b-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="aee4b-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="aee4b-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="aee4b-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="aee4b-165">TabSlideMaster</span></span>
- <span data-ttu-id="aee4b-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="aee4b-166">TabHandoutMaster</span></span>
- <span data-ttu-id="aee4b-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="aee4b-167">TabNotesMaster</span></span>
- <span data-ttu-id="aee4b-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="aee4b-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="aee4b-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="aee4b-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="aee4b-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="aee4b-170">OneNote</span></span>

- <span data-ttu-id="aee4b-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="aee4b-171">**TabHome**</span></span>
- <span data-ttu-id="aee4b-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="aee4b-172">**TabInsert**</span></span>
- <span data-ttu-id="aee4b-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="aee4b-173">**TabView**</span></span>
- <span data-ttu-id="aee4b-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="aee4b-174">TabDeveloper</span></span>
- <span data-ttu-id="aee4b-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="aee4b-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="aee4b-176">Group</span><span class="sxs-lookup"><span data-stu-id="aee4b-176">Group</span></span>

<span data-ttu-id="aee4b-177">Группа точек расширения пользовательского интерфейса на вкладке. У группы может быть до шести элементов управления.</span><span class="sxs-lookup"><span data-stu-id="aee4b-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="aee4b-178">Атрибут **ID** является обязательным, а каждый **идентификатор** должен быть уникальным в пределах манифеста.</span><span class="sxs-lookup"><span data-stu-id="aee4b-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="aee4b-179">**Идентификатор** — это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="aee4b-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="aee4b-180">Просмотрите [элемент Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="aee4b-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="aee4b-181">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="aee4b-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
