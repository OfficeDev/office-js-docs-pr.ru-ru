---
title: Убедитесь, что надстройка Office совместима с существующей надстройкой COM
description: Включить совместимость между Office надстройки и эквивалентной надстройки COM.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: c475189ec19cbb0de1a070f6c529db3c59a093d8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349807"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="30eb5-103">Убедитесь, что надстройка Office совместима с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="30eb5-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="30eb5-104">Если у вас есть существующая надстройка COM, вы можете создать эквивалентную функциональность в Office надстройки, что позволит вашему решению работать на других платформах, таких как Office в Интернете или Mac.</span><span class="sxs-lookup"><span data-stu-id="30eb5-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="30eb5-105">В некоторых случаях Office надстройка может быть не в состоянии предоставить все функциональные возможности, доступные в соответствующей надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="30eb5-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="30eb5-106">В таких ситуациях надстройка COM может предоставлять пользователям более Windows, чем соответствующие Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="30eb5-107">Можно настроить надстройку Office так, чтобы при установке эквивалентной надстройки COM на компьютере пользователя Office на Windows надстройка COM вместо надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="30eb5-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="30eb5-108">Надстройка COM называется "эквивалентной", так как Office плавно переходит между надстройки COM и надстройки Office, в соответствии с которой устанавливается компьютер пользователя.</span><span class="sxs-lookup"><span data-stu-id="30eb5-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="30eb5-109">Эта функция поддерживается следующими платформами при под подключении к Microsoft 365 подписке.</span><span class="sxs-lookup"><span data-stu-id="30eb5-109">This feature is supported by the following platforms, when connected to a Microsoft 365 subscription.</span></span>
>
> - <span data-ttu-id="30eb5-110">Excel, Word и PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="30eb5-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="30eb5-111">Excel, Word и PowerPoint на Windows (версия 1904 или более поздней версии)</span><span class="sxs-lookup"><span data-stu-id="30eb5-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="30eb5-112">Excel, Word и PowerPoint Mac (версия 13.329 или более поздней версии)</span><span class="sxs-lookup"><span data-stu-id="30eb5-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>
> - <span data-ttu-id="30eb5-113">Outlook на Windows (версия 2102 или более поздней версии)</span><span class="sxs-lookup"><span data-stu-id="30eb5-113">Outlook on Windows (version 2102 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="30eb5-114">Укажите эквивалентную надстройка COM</span><span class="sxs-lookup"><span data-stu-id="30eb5-114">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="30eb5-115">Манифест</span><span class="sxs-lookup"><span data-stu-id="30eb5-115">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="30eb5-116">Применяется к Excel, PowerPoint и Word.</span><span class="sxs-lookup"><span data-stu-id="30eb5-116">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="30eb5-117">Outlook поддержка скоро.</span><span class="sxs-lookup"><span data-stu-id="30eb5-117">Outlook support coming soon.</span></span>

<span data-ttu-id="30eb5-118">Чтобы обеспечить совместимость Office надстройки и надстройки COM, определите эквивалентную надстройка COM в манифесте Office надстройки. [](add-in-manifests.md)</span><span class="sxs-lookup"><span data-stu-id="30eb5-118">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="30eb5-119">Затем Office на Windows надстройка COM вместо надстройки Office, если они установлены.</span><span class="sxs-lookup"><span data-stu-id="30eb5-119">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="30eb5-120">В следующем примере показана часть манифеста, которая указывает надстройки COM в качестве эквивалентной надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-120">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="30eb5-121">Значение элемента определяет надстройку COM, а элемент EquivalentAddins должен быть позиционен непосредственно `ProgId` перед закрывающими [](../reference/manifest/equivalentaddins.md) `VersionOverrides` тегами.</span><span class="sxs-lookup"><span data-stu-id="30eb5-121">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="30eb5-122">Сведения о совместимости надстройки COM и совместимости XLL UDF см. в ссылке [Make your custom functions compatible with XLL user-defined functions.](../excel/make-custom-functions-compatible-with-xll-udf.md)</span><span class="sxs-lookup"><span data-stu-id="30eb5-122">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="30eb5-123">Групповая политика</span><span class="sxs-lookup"><span data-stu-id="30eb5-123">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="30eb5-124">Применяется только Outlook.</span><span class="sxs-lookup"><span data-stu-id="30eb5-124">Applies to Outlook only.</span></span>

<span data-ttu-id="30eb5-125">Чтобы объявить совместимость между Outlook веб-надстройки и надстройки COM/VSTO, определите эквивалентную надстройку COM в групповой политике **Deactivate Outlook** веб-надстроек, эквивалентные com или VSTO надстройки, установленные путем настройки на компьютере пользователя.</span><span class="sxs-lookup"><span data-stu-id="30eb5-125">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="30eb5-126">Затем Outlook на Windows будет использовать надстройки COM вместо веб-надстройки, если они установлены.</span><span class="sxs-lookup"><span data-stu-id="30eb5-126">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="30eb5-127">Скачайте последний [инструмент административных шаблонов,](https://www.microsoft.com/download/details.aspx?id=49030)обращая внимание на инструкции по установке **средства.**</span><span class="sxs-lookup"><span data-stu-id="30eb5-127">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="30eb5-128">Откройте редактор локальной групповой политики **(gpedit.msc).**</span><span class="sxs-lookup"><span data-stu-id="30eb5-128">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="30eb5-129">Перейдите **к административным** шаблонам конфигурации  >     >  **пользователей Microsoft Outlook 2016**  >  **разных типов.**</span><span class="sxs-lookup"><span data-stu-id="30eb5-129">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="30eb5-130">Выберите параметр **Deactivate Outlook веб-надстроек,** у которых установлен эквивалент com или VSTO надстройка.</span><span class="sxs-lookup"><span data-stu-id="30eb5-130">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="30eb5-131">Откройте ссылку для редактирования параметра политики.</span><span class="sxs-lookup"><span data-stu-id="30eb5-131">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="30eb5-132">В диалоговом **Outlook веб-надстроек для отключения:**</span><span class="sxs-lookup"><span data-stu-id="30eb5-132">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="30eb5-133">Установите **имя value** для найденного в манифесте `Id` веб-надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-133">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="30eb5-134">**Важно.** *Не добавляйте* фигурные скобки `{}` вокруг входа.</span><span class="sxs-lookup"><span data-stu-id="30eb5-134">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="30eb5-135">**Задайте** значение `ProgId` эквивалентной надстройки COM/VSTO.</span><span class="sxs-lookup"><span data-stu-id="30eb5-135">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="30eb5-136">Выберите **ОК,** чтобы вложить обновление в действие.</span><span class="sxs-lookup"><span data-stu-id="30eb5-136">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="30eb5-137">![Снимок экрана, показывающий диалоговое окно "Outlook веб-надстроек для деактивации".](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="30eb5-137">![Screenshot showing the dialog "Outlook web add-ins to deactivate".](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="30eb5-138">Эквивалентное поведение для пользователей</span><span class="sxs-lookup"><span data-stu-id="30eb5-138">Equivalent behavior for users</span></span>

<span data-ttu-id="30eb5-139">При [указании](#specify-an-equivalent-com-add-in)эквивалентной надстройки COM Office на Windows не будет отображаться пользовательский интерфейс Office надстройки (UI), если установлена эквивалентная надстройка COM.</span><span class="sxs-lookup"><span data-stu-id="30eb5-139">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="30eb5-140">Office только скрывает кнопки ленты надстройки Office надстройки и не препятствует установке.</span><span class="sxs-lookup"><span data-stu-id="30eb5-140">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="30eb5-141">Поэтому Office надстройка по-прежнему будет отображаться в следующих расположениях в пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="30eb5-141">Therefore your Office Add-in will still appear in the following locations within the UI.</span></span>

- <span data-ttu-id="30eb5-142">В **статье Мои надстройки**</span><span class="sxs-lookup"><span data-stu-id="30eb5-142">Under **My add-ins**</span></span>
- <span data-ttu-id="30eb5-143">В качестве записи в диспетчере ленты (только Excel, Word и PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="30eb5-143">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="30eb5-144">Указание эквивалентной надстройки COM в манифесте не влияет на другие платформы, такие как Office в Интернете или Mac.</span><span class="sxs-lookup"><span data-stu-id="30eb5-144">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="30eb5-145">В следующих сценариях описывается, что происходит в зависимости от того, как пользователь Office надстройку.</span><span class="sxs-lookup"><span data-stu-id="30eb5-145">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="30eb5-146">Приобретение appSource Office надстройки</span><span class="sxs-lookup"><span data-stu-id="30eb5-146">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="30eb5-147">Если пользователь приобретает надстройки Office AppSource и эквивалентная надстройка COM уже установлена, Office будет:</span><span class="sxs-lookup"><span data-stu-id="30eb5-147">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="30eb5-148">Установите Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-148">Install the Office Add-in.</span></span>
2. <span data-ttu-id="30eb5-149">Скрыть интерфейс Office надстройки в ленте.</span><span class="sxs-lookup"><span data-stu-id="30eb5-149">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="30eb5-150">Отображение вызова для пользователя, который указывает кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="30eb5-150">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="30eb5-151">Централизованное развертывание Office надстройки</span><span class="sxs-lookup"><span data-stu-id="30eb5-151">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="30eb5-152">Если администратор развертывает надстройку Office клиента с помощью централизованного развертывания, а эквивалентная надстройка COM уже установлена, пользователь должен перезапустить Office, прежде чем они увидят какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="30eb5-152">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="30eb5-153">После Office перезапуска будет:</span><span class="sxs-lookup"><span data-stu-id="30eb5-153">After Office restarts, it will:</span></span>

1. <span data-ttu-id="30eb5-154">Установите Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-154">Install the Office Add-in.</span></span>
2. <span data-ttu-id="30eb5-155">Скрыть интерфейс Office надстройки в ленте.</span><span class="sxs-lookup"><span data-stu-id="30eb5-155">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="30eb5-156">Отображение вызова для пользователя, который указывает кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="30eb5-156">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="30eb5-157">Документ, общий со встроенными Office надстройки</span><span class="sxs-lookup"><span data-stu-id="30eb5-157">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="30eb5-158">Если у пользователя установлена надстройка COM, а затем он получает общий документ со встроенной надстройки Office, то при открываемом документе Office:</span><span class="sxs-lookup"><span data-stu-id="30eb5-158">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="30eb5-159">Запрос пользователя на доверие Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-159">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="30eb5-160">При доверии Office надстройка будет устанавливаться.</span><span class="sxs-lookup"><span data-stu-id="30eb5-160">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="30eb5-161">Скрыть интерфейс Office надстройки в ленте.</span><span class="sxs-lookup"><span data-stu-id="30eb5-161">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="30eb5-162">Другое поведение надстройки COM</span><span class="sxs-lookup"><span data-stu-id="30eb5-162">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="30eb5-163">Excel, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="30eb5-163">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="30eb5-164">Если пользователь отреставрирует эквивалентную надстройку COM, Office на Windows восстанавливает пользовательский интерфейс Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-164">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="30eb5-165">После указания эквивалентной надстройки COM для Office надстройки Office обработку обновлений Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="30eb5-165">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="30eb5-166">Чтобы получить последние обновления для надстройки Office, пользователю необходимо сначала удалить надстройку COM.</span><span class="sxs-lookup"><span data-stu-id="30eb5-166">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="30eb5-167">Outlook</span><span class="sxs-lookup"><span data-stu-id="30eb5-167">Outlook</span></span>

<span data-ttu-id="30eb5-168">Надстройка com/VSTO должна быть подключена при Outlook, чтобы соответствующая веб-надстройка была отключена.</span><span class="sxs-lookup"><span data-stu-id="30eb5-168">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="30eb5-169">Если надстройка com/VSTO отключена во время последующего сеанса Outlook, веб-надстройка, скорее всего, будет отключена до Outlook перезапуска.</span><span class="sxs-lookup"><span data-stu-id="30eb5-169">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="30eb5-170">См. также</span><span class="sxs-lookup"><span data-stu-id="30eb5-170">See also</span></span>

- [<span data-ttu-id="30eb5-171">Совместите пользовательские функции с определенными функциями пользователя XLL</span><span class="sxs-lookup"><span data-stu-id="30eb5-171">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
