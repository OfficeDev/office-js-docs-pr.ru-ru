---
title: Убедитесь, что надстройка Office совместима с существующей надстройкой COM
description: Включение совместимости между надстройкой Office и эквивалентной надстройкой COM
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: ff47b75e8e560bc891c84dc839b7eceffb2400be
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609424"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="13a71-103">Убедитесь, что надстройка Office совместима с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="13a71-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="13a71-104">Если у вас есть надстройка COM, вы можете создавать эквивалентные функциональные возможности в надстройке Office, что позволяет выполнять решение на других платформах, таких как Office в Интернете или Office для Mac.</span><span class="sxs-lookup"><span data-stu-id="13a71-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Office on Mac.</span></span> <span data-ttu-id="13a71-105">В некоторых случаях надстройка Office может не поддерживать все функции, доступные в соответствующей надстройке COM.</span><span class="sxs-lookup"><span data-stu-id="13a71-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="13a71-106">В таких ситуациях надстройка COM может улучшить взаимодействие с пользователем в Windows, а не с соответствующей надстройкой Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="13a71-107">Вы можете настроить надстройку Office таким образом, чтобы когда эквивалентная надстройка COM уже установлена на компьютере пользователя, Office в Windows запускает надстройку COM, а не надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="13a71-108">Надстройка COM называется "эквивалентной", так как Office будет беспрепятственно переходить между надстройкой COM и надстройкой Office в зависимости от того, какой из них установил компьютер пользователя.</span><span class="sxs-lookup"><span data-stu-id="13a71-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="13a71-109">Эта функция поддерживается следующими платформами при подключении к подписке на Office 365:</span><span class="sxs-lookup"><span data-stu-id="13a71-109">This feature is supported by the following platforms, when connected to an Office 365 subscription:</span></span>
> - <span data-ttu-id="13a71-110">Excel, Word и PowerPoint в Интернете</span><span class="sxs-lookup"><span data-stu-id="13a71-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="13a71-111">Excel, Word и PowerPoint для Windows (версия 1904 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="13a71-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="13a71-112">Excel, Word и PowerPoint на Mac (версия 13,329 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="13a71-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="13a71-113">Указание эквивалентной надстройки COM в манифесте</span><span class="sxs-lookup"><span data-stu-id="13a71-113">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="13a71-114">Чтобы обеспечить совместимость надстройки Office и надстройки COM, определите эквивалентную надстройку COM в [манифесте](add-in-manifests.md) надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-114">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="13a71-115">После этого Office в Windows будет использовать надстройку COM, а не надстройку Office, если они установлены.</span><span class="sxs-lookup"><span data-stu-id="13a71-115">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="13a71-116">В следующем примере показана часть манифеста, указывающая надстройку COM в качестве эквивалентной надстройки.</span><span class="sxs-lookup"><span data-stu-id="13a71-116">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="13a71-117">Значение `ProgId` элемента определяет надстройку COM, и `EquivalentAddins` элемент должен быть расположен сразу перед закрывающим `VersionOverrides` тегом.</span><span class="sxs-lookup"><span data-stu-id="13a71-117">The value of the `ProgId` element identifies the COM add-in and the `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

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
> <span data-ttu-id="13a71-118">Сведения о надстройках COM и совместимости с UDF UDF можно найти в статье [Создание пользовательских функций, совместимых с пользовательскими ФУНКЦИЯМИ XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="13a71-118">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="13a71-119">Эквивалентное поведение для пользователей</span><span class="sxs-lookup"><span data-stu-id="13a71-119">Equivalent behavior for users</span></span>

<span data-ttu-id="13a71-120">Если в манифесте надстройки Office указана эквивалентная надстройка COM, Office в Windows не будет отображать пользовательский интерфейс надстройки Office, если установлена эквивалентная надстройка COM.</span><span class="sxs-lookup"><span data-stu-id="13a71-120">When an equivalent COM add-in is specified in the Office Add-in manifest, Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="13a71-121">Office скрывает кнопки ленты только в надстройке Office и не запрещает установку.</span><span class="sxs-lookup"><span data-stu-id="13a71-121">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="13a71-122">Поэтому надстройка Office будет по-прежнему отображаться в следующих расположениях в пользовательском интерфейсе:</span><span class="sxs-lookup"><span data-stu-id="13a71-122">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="13a71-123">В разделе **Мои надстройки**</span><span class="sxs-lookup"><span data-stu-id="13a71-123">Under **My add-ins**</span></span>
- <span data-ttu-id="13a71-124">Как запись в диспетчере лент</span><span class="sxs-lookup"><span data-stu-id="13a71-124">As an entry in the ribbon manager</span></span>

> [!NOTE]
> <span data-ttu-id="13a71-125">Указание эквивалентной надстройки COM в манифесте не оказывает никакого действия на других платформах, таких как Office в Интернете или Mac.</span><span class="sxs-lookup"><span data-stu-id="13a71-125">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or Mac.</span></span>

<span data-ttu-id="13a71-126">В следующих сценариях описывается, что происходит в зависимости от того, как пользователь приобретает надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-126">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="13a71-127">AppSource приобретение надстройки Office</span><span class="sxs-lookup"><span data-stu-id="13a71-127">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="13a71-128">Если пользователь приобретает надстройку Office из AppSource, а эквивалентная надстройка COM уже установлена, Office выполнит следующие действия:</span><span class="sxs-lookup"><span data-stu-id="13a71-128">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="13a71-129">Установите надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-129">Install the Office Add-in.</span></span>
2. <span data-ttu-id="13a71-130">Скрытие пользовательского интерфейса надстройки Office на ленте.</span><span class="sxs-lookup"><span data-stu-id="13a71-130">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="13a71-131">Отображение вызываемого абонента для пользователя, который указывает на кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="13a71-131">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="13a71-132">Централизованное развертывание надстройки Office</span><span class="sxs-lookup"><span data-stu-id="13a71-132">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="13a71-133">Если Администратор развертывает надстройку Office в своем клиенте с помощью централизованного развертывания, а эквивалентная надстройка COM уже установлена, пользователь должен перезапустить Office, чтобы увидеть изменения.</span><span class="sxs-lookup"><span data-stu-id="13a71-133">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="13a71-134">После перезапуска Office будет:</span><span class="sxs-lookup"><span data-stu-id="13a71-134">After Office restarts, it will:</span></span>

1. <span data-ttu-id="13a71-135">Установите надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-135">Install the Office Add-in.</span></span>
2. <span data-ttu-id="13a71-136">Скрытие пользовательского интерфейса надстройки Office на ленте.</span><span class="sxs-lookup"><span data-stu-id="13a71-136">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="13a71-137">Отображение вызываемого абонента для пользователя, который указывает на кнопку ленты надстройки COM.</span><span class="sxs-lookup"><span data-stu-id="13a71-137">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="13a71-138">Общий доступ к документу с помощью встроенной надстройки Office</span><span class="sxs-lookup"><span data-stu-id="13a71-138">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="13a71-139">Если у пользователя установлена надстройка COM, а затем он получает общий документ с внедренной надстройкой Office, то при открытии документа Office будет:</span><span class="sxs-lookup"><span data-stu-id="13a71-139">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="13a71-140">Предложит пользователю доверять надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-140">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="13a71-141">Если вы доверяете, надстройка Office будет установлена.</span><span class="sxs-lookup"><span data-stu-id="13a71-141">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="13a71-142">Скрытие пользовательского интерфейса надстройки Office на ленте.</span><span class="sxs-lookup"><span data-stu-id="13a71-142">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="13a71-143">Другое поведение надстройки COM</span><span class="sxs-lookup"><span data-stu-id="13a71-143">Other COM add-in behavior</span></span>

<span data-ttu-id="13a71-144">Если пользователь удаляет эквивалентную надстройку COM, Office в Windows восстанавливает пользовательский интерфейс надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-144">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="13a71-145">После того как вы укажете эквивалентную надстройку COM для надстройки Office, Office прекратит обработку обновлений для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="13a71-145">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="13a71-146">Чтобы получить последние обновления для надстройки Office, пользователь должен сначала удалить надстройку COM.</span><span class="sxs-lookup"><span data-stu-id="13a71-146">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="13a71-147">См. также</span><span class="sxs-lookup"><span data-stu-id="13a71-147">See also</span></span>

- [<span data-ttu-id="13a71-148">Обеспечение совместимости пользовательских функций с пользовательскими функциями XLL</span><span class="sxs-lookup"><span data-stu-id="13a71-148">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
