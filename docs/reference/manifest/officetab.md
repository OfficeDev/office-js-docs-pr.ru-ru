---
title: Элемент OfficeTab в файле манифеста
description: Элемент OfficeTab определяет вкладку ленты, в которой отображается команда надстройки.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9b07ce1e57329e796545610e0c61a2c11d1ed55d
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641443"
---
# <a name="officetab-element"></a><span data-ttu-id="bcfad-103">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bcfad-103">OfficeTab element</span></span>

<span data-ttu-id="bcfad-104">Определяет вкладку ленты, на которой отображается команда надстройки.</span><span class="sxs-lookup"><span data-stu-id="bcfad-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="bcfad-105">Это может быть вкладка по умолчанию (" **домашний**", " **сообщение**" или " **собрание**") или настраиваемая вкладка, определенная надстройкой.</span><span class="sxs-lookup"><span data-stu-id="bcfad-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="bcfad-106">Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="bcfad-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bcfad-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="bcfad-107">Child elements</span></span>

|  <span data-ttu-id="bcfad-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="bcfad-108">Element</span></span> |  <span data-ttu-id="bcfad-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bcfad-109">Required</span></span>  |  <span data-ttu-id="bcfad-110">Описание</span><span class="sxs-lookup"><span data-stu-id="bcfad-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bcfad-111">Группа</span><span class="sxs-lookup"><span data-stu-id="bcfad-111">Group</span></span>      | <span data-ttu-id="bcfad-112">Да</span><span class="sxs-lookup"><span data-stu-id="bcfad-112">Yes</span></span> |  <span data-ttu-id="bcfad-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="bcfad-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="bcfad-115">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="bcfad-115">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="bcfad-116">Значения, **выделенные полужирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word в Интернете).</span><span class="sxs-lookup"><span data-stu-id="bcfad-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="bcfad-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="bcfad-117">Outlook</span></span>

- <span data-ttu-id="bcfad-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="bcfad-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="bcfad-119">Word</span><span class="sxs-lookup"><span data-stu-id="bcfad-119">Word</span></span>

- <span data-ttu-id="bcfad-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcfad-120">**TabHome**</span></span>
- <span data-ttu-id="bcfad-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcfad-121">**TabInsert**</span></span>
- <span data-ttu-id="bcfad-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="bcfad-122">TabWordDesign</span></span>
- <span data-ttu-id="bcfad-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="bcfad-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="bcfad-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="bcfad-124">TabReferences</span></span>
- <span data-ttu-id="bcfad-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="bcfad-125">TabMailings</span></span>
- <span data-ttu-id="bcfad-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="bcfad-126">TabReviewWord</span></span>
- <span data-ttu-id="bcfad-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcfad-127">**TabView**</span></span>
- <span data-ttu-id="bcfad-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcfad-128">TabDeveloper</span></span>
- <span data-ttu-id="bcfad-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcfad-129">TabAddIns</span></span>
- <span data-ttu-id="bcfad-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="bcfad-130">TabBlogPost</span></span>
- <span data-ttu-id="bcfad-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="bcfad-131">TabBlogInsert</span></span>
- <span data-ttu-id="bcfad-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bcfad-132">TabPrintPreview</span></span>
- <span data-ttu-id="bcfad-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="bcfad-133">TabOutlining</span></span>
- <span data-ttu-id="bcfad-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="bcfad-134">TabConflicts</span></span>
- <span data-ttu-id="bcfad-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bcfad-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bcfad-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bcfad-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="bcfad-137">Excel</span><span class="sxs-lookup"><span data-stu-id="bcfad-137">Excel</span></span>

- <span data-ttu-id="bcfad-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcfad-138">**TabHome**</span></span>
- <span data-ttu-id="bcfad-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcfad-139">**TabInsert**</span></span>
- <span data-ttu-id="bcfad-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="bcfad-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="bcfad-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="bcfad-141">TabFormulas</span></span>
- <span data-ttu-id="bcfad-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="bcfad-142">**TabData**</span></span>
- <span data-ttu-id="bcfad-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="bcfad-143">**TabReview**</span></span>
- <span data-ttu-id="bcfad-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcfad-144">**TabView**</span></span>
- <span data-ttu-id="bcfad-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcfad-145">TabDeveloper</span></span>
- <span data-ttu-id="bcfad-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcfad-146">TabAddIns</span></span>
- <span data-ttu-id="bcfad-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bcfad-147">TabPrintPreview</span></span>
- <span data-ttu-id="bcfad-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bcfad-148">TabBackgroundRemoval</span></span>

### <a name="powerpoint"></a><span data-ttu-id="bcfad-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bcfad-149">PowerPoint</span></span>

- <span data-ttu-id="bcfad-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcfad-150">**TabHome**</span></span>
- <span data-ttu-id="bcfad-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcfad-151">**TabInsert**</span></span>
- <span data-ttu-id="bcfad-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="bcfad-152">**TabDesign**</span></span>
- <span data-ttu-id="bcfad-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="bcfad-153">**TabTransitions**</span></span>
- <span data-ttu-id="bcfad-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="bcfad-154">**TabAnimations**</span></span>
- <span data-ttu-id="bcfad-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="bcfad-155">TabSlideShow</span></span>
- <span data-ttu-id="bcfad-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="bcfad-156">TabReview</span></span>
- <span data-ttu-id="bcfad-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcfad-157">**TabView**</span></span>
- <span data-ttu-id="bcfad-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcfad-158">TabDeveloper</span></span>
- <span data-ttu-id="bcfad-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcfad-159">TabAddIns</span></span>
- <span data-ttu-id="bcfad-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bcfad-160">TabPrintPreview</span></span>
- <span data-ttu-id="bcfad-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="bcfad-161">TabMerge</span></span>
- <span data-ttu-id="bcfad-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="bcfad-162">TabGrayscale</span></span>
- <span data-ttu-id="bcfad-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="bcfad-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="bcfad-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bcfad-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="bcfad-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="bcfad-165">TabSlideMaster</span></span>
- <span data-ttu-id="bcfad-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="bcfad-166">TabHandoutMaster</span></span>
- <span data-ttu-id="bcfad-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="bcfad-167">TabNotesMaster</span></span>
- <span data-ttu-id="bcfad-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bcfad-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bcfad-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="bcfad-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="bcfad-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="bcfad-170">OneNote</span></span>

- <span data-ttu-id="bcfad-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcfad-171">**TabHome**</span></span>
- <span data-ttu-id="bcfad-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcfad-172">**TabInsert**</span></span>
- <span data-ttu-id="bcfad-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcfad-173">**TabView**</span></span>
- <span data-ttu-id="bcfad-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcfad-174">TabDeveloper</span></span>
- <span data-ttu-id="bcfad-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcfad-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="bcfad-176">Group</span><span class="sxs-lookup"><span data-stu-id="bcfad-176">Group</span></span>

<span data-ttu-id="bcfad-177">Группа точек расширения пользовательского интерфейса на вкладке. У группы может быть до шести элементов управления.</span><span class="sxs-lookup"><span data-stu-id="bcfad-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="bcfad-178">Атрибут **ID** является обязательным, а каждый **идентификатор** должен быть уникальным в пределах манифеста.</span><span class="sxs-lookup"><span data-stu-id="bcfad-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="bcfad-179">**Идентификатор** — это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="bcfad-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="bcfad-180">Просмотрите [элемент Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="bcfad-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="bcfad-181">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bcfad-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
