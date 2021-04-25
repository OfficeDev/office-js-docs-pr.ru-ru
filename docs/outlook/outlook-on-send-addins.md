---
title: Функция проверки при отправке для надстроек Outlook
description: Позволяет надстройке настраивать те или иные параметры при отправке, а также обрабатывать элемент и запрещать пользователям выполнять определенные действия.
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: 126323527d74553aa7fd7e0c8cf1e5e5d89471ff
ms.sourcegitcommit: 691fa338029c9cbd9a7194d163f390c3321a0cd8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/23/2021
ms.locfileid: "51959182"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="2f0bd-103">Функция проверки при отправке для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="2f0bd-p101">Функция проверки при отправке позволяет надстройкам Outlook настраивать те или иные параметры при отправке, а также обрабатывать сообщение или элемент собрания и запрещать пользователям выполнять определенные действия. Например, с помощью функции проверки при отправке можно сделать следующее:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="2f0bd-106">запретить пользователю отправлять конфиденциальную информацию или оставлять строку темы пустой;</span><span class="sxs-lookup"><span data-stu-id="2f0bd-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="2f0bd-107">добавить определенного получателя в строку "Копия" в сообщениях или в строку "Необязательные получатели" в собраниях.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="2f0bd-108">Функция проверки при отправке не использует пользовательский интерфейс и активируется событием типа `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="2f0bd-109">Сведения об ограничениях, связанных с функцией проверки при отправке, см. в разделе [Ограничения](#limitations) далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="2f0bd-110">Поддерживаемые клиенты и платформы</span><span class="sxs-lookup"><span data-stu-id="2f0bd-110">Supported clients and platforms</span></span>

<span data-ttu-id="2f0bd-111">В следующей таблице показаны поддерживаемые клиенто-серверные комбинации для функции отправки, включая минимально необходимое накопительное обновление, если это применимо.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-111">The following table shows supported client-server combinations for the on-send feature, including the minimum required Cumulative Update where applicable.</span></span> <span data-ttu-id="2f0bd-112">Исключенные комбинации не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="2f0bd-113">Клиент</span><span class="sxs-lookup"><span data-stu-id="2f0bd-113">Client</span></span> | <span data-ttu-id="2f0bd-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="2f0bd-114">Exchange Online</span></span> | <span data-ttu-id="2f0bd-115">Локальное exchange 2016</span><span class="sxs-lookup"><span data-stu-id="2f0bd-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="2f0bd-116">(Накопительное обновление 6 или более позднее)</span><span class="sxs-lookup"><span data-stu-id="2f0bd-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="2f0bd-117">Локальное exchange 2019</span><span class="sxs-lookup"><span data-stu-id="2f0bd-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="2f0bd-118">(Накопительное обновление 1 или более позднее)</span><span class="sxs-lookup"><span data-stu-id="2f0bd-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="2f0bd-119">Windows:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-119">Windows:</span></span><br><span data-ttu-id="2f0bd-120">версия 1910 (сборка 12130.20272) или более поздней версии</span><span class="sxs-lookup"><span data-stu-id="2f0bd-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="2f0bd-121">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-121">Yes</span></span>|<span data-ttu-id="2f0bd-122">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-122">Yes</span></span>|<span data-ttu-id="2f0bd-123">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-123">Yes</span></span>|
|<span data-ttu-id="2f0bd-124">Mac:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-124">Mac:</span></span><br><span data-ttu-id="2f0bd-125">сборка с 16.30 до 16.46</span><span class="sxs-lookup"><span data-stu-id="2f0bd-125">build 16.30 to 16.46</span></span>|<span data-ttu-id="2f0bd-126">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-126">Yes</span></span>|<span data-ttu-id="2f0bd-127">Нет</span><span class="sxs-lookup"><span data-stu-id="2f0bd-127">No</span></span>|<span data-ttu-id="2f0bd-128">Нет</span><span class="sxs-lookup"><span data-stu-id="2f0bd-128">No</span></span>|
|<span data-ttu-id="2f0bd-129">Mac:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-129">Mac:</span></span><br><span data-ttu-id="2f0bd-130">сборка 16.47 или более поздней</span><span class="sxs-lookup"><span data-stu-id="2f0bd-130">build 16.47 or later</span></span>|<span data-ttu-id="2f0bd-131">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-131">Yes</span></span>|<span data-ttu-id="2f0bd-132">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-132">Yes</span></span>|<span data-ttu-id="2f0bd-133">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-133">Yes</span></span>|
|<span data-ttu-id="2f0bd-134">Веб-браузер:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-134">Web browser:</span></span><br><span data-ttu-id="2f0bd-135">современный пользовательский интерфейс Outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-135">modern Outlook UI</span></span>|<span data-ttu-id="2f0bd-136">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-136">Yes</span></span>|<span data-ttu-id="2f0bd-137">Неприменимо</span><span class="sxs-lookup"><span data-stu-id="2f0bd-137">Not applicable</span></span>|<span data-ttu-id="2f0bd-138">Неприменимо</span><span class="sxs-lookup"><span data-stu-id="2f0bd-138">Not applicable</span></span>|
|<span data-ttu-id="2f0bd-139">Веб-браузер:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-139">Web browser:</span></span><br><span data-ttu-id="2f0bd-140">классический пользовательский интерфейс Outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-140">classic Outlook UI</span></span>|<span data-ttu-id="2f0bd-141">Неприменимо</span><span class="sxs-lookup"><span data-stu-id="2f0bd-141">Not applicable</span></span>|<span data-ttu-id="2f0bd-142">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-142">Yes</span></span>|<span data-ttu-id="2f0bd-143">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-143">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="2f0bd-144">Функция отправки была официально выпущена в наборе требований 1.8 (см. сведения о текущем сервере [и клиентской](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) поддержке).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-144">The on-send feature was officially released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span> <span data-ttu-id="2f0bd-145">Однако обратите внимание, что матрица поддержки функции является суперсетью набора требований.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-145">However, note that the feature's support matrix is a superset of the requirement set's.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f0bd-146">Надстройки, которые используют функцию отправки, не допускаются в [AppSource.](https://appsource.microsoft.com)</span><span class="sxs-lookup"><span data-stu-id="2f0bd-146">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="2f0bd-147">Как работает функция проверки при отправке?</span><span class="sxs-lookup"><span data-stu-id="2f0bd-147">How does the on-send feature work?</span></span>

<span data-ttu-id="2f0bd-148">С помощью функции проверки при отправке вы можете создать надстройку Outlook, задействующую синхронное событие `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-148">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="2f0bd-149">Это событие возникает, когда пользователь нажимает кнопку **Отправить** (или **Отправить обновление** для существующих собраний). С его помощью можно блокировать отправку элемента, не прошедшего проверку.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-149">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="2f0bd-150">Например, когда пользователь вызывает событие отправки сообщения, надстройка Outlook благодаря функции проверки при отправке может следующее:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-150">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="2f0bd-151">прочитать и проверить содержимое письма;</span><span class="sxs-lookup"><span data-stu-id="2f0bd-151">Read and validate the email message contents</span></span>
- <span data-ttu-id="2f0bd-152">проверить, содержит ли сообщение строку темы;</span><span class="sxs-lookup"><span data-stu-id="2f0bd-152">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="2f0bd-153">задать заранее определенного получателя.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-153">Set a predetermined recipient</span></span>

<span data-ttu-id="2f0bd-154">Проверка проводится на стороне клиента в Outlook, когда событие отправки запускается, и надстройка имеет до 5 минут, прежде чем он раз. Если проверка не удается, отправка элемента блокируется, а сообщение об ошибке отображается в информационной панели, которая побуждает пользователя к действию.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-154">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

> [!NOTE]
> <span data-ttu-id="2f0bd-155">В Outlook в Интернете, когда функция отправки запускается в сообщении, сочиняемом в вкладке браузера Outlook, элемент выскочил в собственное окно браузера или вкладку, чтобы завершить проверку и другую обработку.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-155">In Outlook on the web, when the on-send feature is triggered in a message being composed within the Outlook browser tab, the item is popped out to its own browser window or tab in order to complete validation and other processing.</span></span>

<span data-ttu-id="2f0bd-156">На приведенном ниже снимке экрана показана панель информации, в которой пользователю предлагается добавить тему.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-156">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![Снимок экрана с сообщением об ошибке, в котором пользователю предлагается ввести отсутствующую строку темы](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="2f0bd-158">На приведенном ниже снимке экрана показана панель информации, уведомляющая отправителя о том, что обнаружены слова, подлежащие блокировке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-158">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![Снимок экрана с сообщением об ошибке, которое сообщает пользователю о том, что обнаружены слова, подлежащие блокировке](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="2f0bd-160">Ограничения</span><span class="sxs-lookup"><span data-stu-id="2f0bd-160">Limitations</span></span>

<span data-ttu-id="2f0bd-161">В настоящее время на функцию проверки при отправке действуют перечисленные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-161">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="2f0bd-162">**Функция append-on-send** &ndash; При вызове [тела. В обработнике отправки appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-162">**Append-on-send** feature &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="2f0bd-163">**AppSource**. Надстройки Outlook, в которых используется функция проверки при отправке, невозможно публиковать в [AppSource](https://appsource.microsoft.com), так как они не проходят проверку AppSource.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-163">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="2f0bd-164">Надстройки, использующие функцию проверки при отправке, должны разворачиваться администраторами.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-164">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="2f0bd-165">**Манифест**. Для каждой надстройки поддерживается только одно событие `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-165">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="2f0bd-166">Если манифест содержит несколько событий `ItemSend`, он не пройдет проверку.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-166">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="2f0bd-p107">**Производительность**. &ndash;Многочисленные случаи приема-передачи пакетов на веб-сервере, где размещается надстройка, могут повлиять на ее производительность. Учитывайте влияние на производительность при создании надстроек, требующих выполнения нескольких операций с сообщениями или собраниями.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="2f0bd-169">**Отправить позже** (только для Mac). Если имеются надстройки, использующие проверку при отправке, функция **Отправить позже** будет недоступна.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-169">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

<span data-ttu-id="2f0bd-170">Кроме того, не рекомендуется вызывать обработитель событий при отправке, так как закрытие элемента должно происходить автоматически после `item.close()` завершения события.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-170">Also, it's not recommended that you call `item.close()` in the on-send event handler as closing the item should happen automatically after the event is completed.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="2f0bd-171">Ограничения на типы и режимы почтовых ящиков</span><span class="sxs-lookup"><span data-stu-id="2f0bd-171">Mailbox type/mode limitations</span></span>

<span data-ttu-id="2f0bd-172">Функция проверки при отправке поддерживается только для почтовых ящиков пользователей в Outlook в Интернете, Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-172">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="2f0bd-173">В настоящее время она не поддерживается для указанных ниже типов и режимов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-173">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="2f0bd-174">Общие почтовые ящики\*</span><span class="sxs-lookup"><span data-stu-id="2f0bd-174">Shared mailboxes\*</span></span>
- <span data-ttu-id="2f0bd-175">почтовые ящики групп;</span><span class="sxs-lookup"><span data-stu-id="2f0bd-175">Group mailboxes</span></span>
- <span data-ttu-id="2f0bd-176">почтовые ящики в автономном режиме.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-176">Offline mode</span></span>

<span data-ttu-id="2f0bd-177">Outlook не разрешает отправку, если для этих сценариев почтовых ящиков включена функция проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-177">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="2f0bd-178">Однако если пользователь отвечает на сообщение в почтовом ящике группы, надстройка, поддерживающая проверку сообщений при отправке, не запускается, а сообщение отправляется.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-178">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f0bd-179">\*Функции отправки должны работать на общих почтовых ящиках или папках, если надстройка также реализует поддержку сценариев доступа [к делегатам.](delegate-access.md)</span><span class="sxs-lookup"><span data-stu-id="2f0bd-179">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="2f0bd-180">Несколько надстроек, поддерживающих проверку сообщений при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-180">Multiple on-send add-ins</span></span>

<span data-ttu-id="2f0bd-181">Если установлено несколько надстроек, поддерживающих проверку сообщений при отправке, они будут запускаться в том порядке, в котором были получены из API `getAppManifestCall` или `getExtensibilityContext`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-181">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="2f0bd-182">Если первая надстройка разрешает отправку сообщения, то вторая может внести изменения, на которые первая отреагировала бы запретом отправки.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-182">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="2f0bd-183">Однако первая надстройка не будет запускаться снова, если все установленные надстройки разрешат отправку.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-183">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="2f0bd-184">Допустим, надстройка 1 и надстройка 2 используют функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-184">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="2f0bd-185">Сначала устанавливается надстройка 1, а затем — надстройка 2.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-185">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="2f0bd-186">Надстройка 1 находит в сообщении слово Fabrikam, что является условием для разрешения отправки.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-186">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="2f0bd-187">Однако надстройка 2 удаляет все вхождения слова Fabrikam.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-187">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="2f0bd-188">Сообщение будет отправлено без единого слова Fabrikam (в связи с порядком установки надстроек 1 и 2).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-188">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="2f0bd-189">Развертывание надстроек Outlook, использующих функцию проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-189">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="2f0bd-190">Рекомендуем, чтобы развертывание надстроек Outlook, использующих функцию проверки при отправке, выполняли администраторы.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-190">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="2f0bd-191">Администратор должен убедиться, что такая надстройка:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-191">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="2f0bd-192">всегда присутствует при открытии создаваемого элемента (для электронной почты: создание сообщений, ответ и пересылка);</span><span class="sxs-lookup"><span data-stu-id="2f0bd-192">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="2f0bd-193">не может быть закрыта или отключена пользователем.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-193">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="2f0bd-194">Установка надстроек Outlook, использующих функцию проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-194">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="2f0bd-195">Чтобы использовать функцию проверки при отправке в Outlook, надстройки должны быть настроены для типов событий отправки.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-195">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="2f0bd-196">Выберите платформу, которую нужно настроить.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-196">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="2f0bd-197">Веб-браузер — классическая версия Outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-197">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="2f0bd-198">Надстройки для Outlook в Интернете (классическая версия), использующие функцию проверки при отправке, будут запускаться у пользователей, которым назначена политика почтовых ящиков Outlook в Интернете, для флага *OnSendAddinsEnabled* которой задано значение **true**.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-198">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="2f0bd-199">Чтобы установить новую надстройку, выполните приведенные ниже командлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-199">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="2f0bd-200">Сведения о том, как подключиться к Exchange Online с помощью удаленного сеанса PowerShell, см. в статье [Подключение к Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-200">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="2f0bd-201">Включение функции проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-201">Enable the on-send feature</span></span>

<span data-ttu-id="2f0bd-202">По умолчанию функция проверки при отправке отключена.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-202">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="2f0bd-203">Администраторы могут включать эту функцию с помощью командлетов Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-203">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="2f0bd-204">Чтобы для всех пользователей включить надстройки, поддерживающие проверку сообщений при отправке, сделайте следующее:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-204">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="2f0bd-205">Создайте политику почтовых ящиков Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-205">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="2f0bd-206">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-206">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="2f0bd-207">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-207">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="2f0bd-208">Включите функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-208">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="2f0bd-209">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-209">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="2f0bd-210">Включение функции проверки при отправке для группы пользователей</span><span class="sxs-lookup"><span data-stu-id="2f0bd-210">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="2f0bd-211">Чтобы включить функцию проверки при отправке для определенной группы пользователей, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-211">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="2f0bd-212">В этом примере администратор включает функцию проверки при отправке для надстроек Outlook в Интернете только в среде финансового отдела (Finance).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-212">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="2f0bd-213">Создайте политику почтовых ящиков Outlook в Интернете для группы.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-213">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="2f0bd-214">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков (дополнительные сведения см. в разделе [Ограничения на типы почтовых ящиков](#multiple-on-send-add-ins) выше в этой статье).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-214">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="2f0bd-215">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-215">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="2f0bd-216">Включите функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-216">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="2f0bd-217">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-217">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="2f0bd-218">Дождитесь вступления политики в силу (это может занять до 60 минут) или перезапустите службы IIS.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-218">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="2f0bd-219">Когда политика вступит в силу, для группы будет включена функция проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-219">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="2f0bd-220">Отключение функции проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-220">Disable the on-send feature</span></span>

<span data-ttu-id="2f0bd-221">Чтобы отключить функцию проверки при отправке для пользователя или назначить политику почтовых ящиков Outlook в Интернете, в которой не включен соответствующий флаг, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-221">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="2f0bd-222">В этом примере используется политика почтовых ящиков *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-222">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="2f0bd-223">Дополнительные сведения о том, как настроить существующие политики почтовых ящиков Outlook в Интернете с помощью командлета **Set-OwaMailboxPolicy**, см. в статье [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-223">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="2f0bd-224">Чтобы отключить функцию проверки при отправке для всех пользователей, которым назначена определенная политика почтовых ящиков Outlook в Интернете, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-224">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="2f0bd-225">Веб-браузер — современная версия Outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-225">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="2f0bd-226">Надстройки для Outlook в Интернете (современная версия), использующие функцию проверки при отправке, должны запускаться для всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-226">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="2f0bd-227">Однако если пользователям необходимо запускать надстройки по отправке в соответствии со стандартами соответствия требованиям, политика почтовых ящиков должна иметь флаг *OnSendAddinsEnabled,* чтобы редактирование элемента не было разрешено во время обработки надстройок при `true` отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-227">However, if users are required to run on-send add-ins to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to `true` so that editing the item is not allowed while the add-ins are processing on send.</span></span>

<span data-ttu-id="2f0bd-228">Чтобы установить новую надстройку, выполните приведенные ниже командлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-228">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="2f0bd-229">Сведения о том, как подключиться к Exchange Online с помощью удаленного сеанса PowerShell, см. в статье [Подключение к Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-229">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-flag"></a><span data-ttu-id="2f0bd-230">Включить флаг отправки</span><span class="sxs-lookup"><span data-stu-id="2f0bd-230">Enable the on-send flag</span></span>

<span data-ttu-id="2f0bd-231">Администраторы могут обеспечивать соответствие требованиям при отправке, запуская комдлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-231">Administrators can enforce on-send compliance by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="2f0bd-232">Чтобы отослать редактирование во время обработки надстройок при отправке, для всех пользователей необходимо отослать:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-232">For all users, to disallow editing while on-send add-ins are processing:</span></span>

1. <span data-ttu-id="2f0bd-233">Создайте политику почтовых ящиков Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-233">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="2f0bd-234">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-234">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="2f0bd-235">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-235">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="2f0bd-236">Обеспечение соответствия требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-236">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="2f0bd-237">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-237">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a><span data-ttu-id="2f0bd-238">Включив флаг отправки для группы пользователей</span><span class="sxs-lookup"><span data-stu-id="2f0bd-238">Turn on the on-send flag for a group of users</span></span>

<span data-ttu-id="2f0bd-239">Чтобы обеспечить соблюдение при отправке для определенной группы пользователей, выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-239">To enforce on-send compliance for a specific group of users, the steps are as follows.</span></span> <span data-ttu-id="2f0bd-240">В этом примере администратор включает политику проверки при отправке для надстроек Outlook в Интернете только в среде финансового отдела (Finance).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-240">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="2f0bd-241">Создайте политику почтовых ящиков Outlook в Интернете для группы.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-241">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="2f0bd-242">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков (дополнительные сведения см. в разделе [Ограничения на типы почтовых ящиков](#multiple-on-send-add-ins) выше в этой статье).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-242">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="2f0bd-243">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-243">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="2f0bd-244">Обеспечение соответствия требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-244">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="2f0bd-245">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-245">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="2f0bd-246">Дождитесь вступления политики в силу (это может занять до 60 минут) или перезапустите службы IIS.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-246">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="2f0bd-247">Когда политика вступает в силу, для группы будет применено соответствие требованиям по отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-247">When the policy takes effect, on-send compliance will be enforced for the group.</span></span>

#### <a name="turn-off-the-on-send-flag"></a><span data-ttu-id="2f0bd-248">Выключите флаг отправки</span><span class="sxs-lookup"><span data-stu-id="2f0bd-248">Turn off the on-send flag</span></span>

<span data-ttu-id="2f0bd-249">Чтобы отключить для пользователя правоприменительных требований к выполнению при отправке, назначьте Outlook в политике веб-почтовых ящиков, в которую не включен флаг, запуская следующие cmdlets.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-249">To turn off on-send compliance enforcement for a user, assign an Outlook on the web mailbox policy that does not have the flag enabled by running the following cmdlets.</span></span> <span data-ttu-id="2f0bd-250">В этом примере используется политика почтовых ящиков *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-250">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="2f0bd-251">Дополнительные сведения о том, как настроить существующие политики почтовых ящиков Outlook в Интернете с помощью командлета **Set-OwaMailboxPolicy**, см. в статье [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-251">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="2f0bd-252">Чтобы отключить правоприменительных требований для всех пользователей с определенным Outlook в назначенной политике почтовых ящиков, запустите следующие cmdlets.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-252">To turn off on-send compliance enforcement for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="2f0bd-253">Windows</span><span class="sxs-lookup"><span data-stu-id="2f0bd-253">Windows</span></span>](#tab/windows)

<span data-ttu-id="2f0bd-254">Надстройки Outlook для Windows, использующие функцию проверки при отправке, должны запускаться для всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-254">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="2f0bd-255">Однако если пользователям требуется запустить надстройку для выполнения стандартов соответствия требованиям, групповой политике **Отключить отправку, если загрузка веб-расширений невозможна** необходимо присвоить значение **Включено** на каждом применяемом компьютере.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-255">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="2f0bd-256">Чтобы установить политики почтовых ящиков, [](https://www.microsoft.com/download/details.aspx?id=49030) администраторы могут скачать средство административных шаблонов, а затем получить доступ к последним шаблонам администрирования, заведя редактор локальной групповой политики **gpedit.msc**.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-256">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy Editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="2f0bd-257">Действия политики</span><span class="sxs-lookup"><span data-stu-id="2f0bd-257">What the policy does</span></span>

<span data-ttu-id="2f0bd-258">Для соответствия требованиям администраторам может потребоваться отключить возможность отправки сообщения или элементов собрания пользователями, пока не станет доступна для запуска последняя версия надстройки с функцией проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-258">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="2f0bd-259">Администраторы должны включить групповую политику **Отключить отправку, если загрузка веб-расширений невозможна**, чтобы все надстройки обновлялись из службы Exchange и были доступны для проверки того, что каждое сообщение или элемент собрания соответствует ожидаемым правилам и нормативным требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-259">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="2f0bd-260">Состояние политики</span><span class="sxs-lookup"><span data-stu-id="2f0bd-260">Policy status</span></span>|<span data-ttu-id="2f0bd-261">Результат</span><span class="sxs-lookup"><span data-stu-id="2f0bd-261">Result</span></span>|
|---|---|
|<span data-ttu-id="2f0bd-262">Отключено</span><span class="sxs-lookup"><span data-stu-id="2f0bd-262">Disabled</span></span>|<span data-ttu-id="2f0bd-263">Загруженные в настоящее время манифесты надстройок (не обязательно последних версий) запускаются при отправке сообщений или элементов собраний.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-263">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="2f0bd-264">Это состояние или поведение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-264">This is the default status/behavior.</span></span>|
|<span data-ttu-id="2f0bd-265">Включен</span><span class="sxs-lookup"><span data-stu-id="2f0bd-265">Enabled</span></span>|<span data-ttu-id="2f0bd-266">После скачивания последних манифестов надстройки по отправке из Exchange надстройки запускаются на отправленных сообщениях или собраниях.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-266">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="2f0bd-267">В противном случае отправка блокируется.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-267">Otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="2f0bd-268">Управление политикой проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-268">Manage the on-send policy</span></span>

<span data-ttu-id="2f0bd-269">По умолчанию политика проверки при отправке отключена.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-269">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="2f0bd-270">Администраторы могут включить политику проверки при отправке, присвоив параметру групповой политики пользователя **Отключить отправку, если загрузка веб-расширений невозможна** значение **Включено**.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-270">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="2f0bd-271">Чтобы отключить политику для пользователя, администратору следует присвоить ей значение **Отключено**.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-271">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="2f0bd-272">Чтобы управлять этим параметром политики, можно сделать следующее:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-272">To manage this policy setting, you can do the following:</span></span>

1. <span data-ttu-id="2f0bd-273">Скачайте последнее [средство административных шаблонов](https://www.microsoft.com/download/details.aspx?id=49030).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-273">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="2f0bd-274">Откройте редактор локальной групповой политики **(gpedit.msc).**</span><span class="sxs-lookup"><span data-stu-id="2f0bd-274">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="2f0bd-275">Выберите **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-275">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="2f0bd-276">Выберите параметр **Отключить отправку, если загрузка веб-расширений невозможна**.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-276">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="2f0bd-277">Откройте ссылку для изменения параметра политики.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-277">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="2f0bd-278">В диалоговом окне **Отключить отправку, если загрузка веб-расширений невозможна** выберите нужный параметр (**Включено** или **Отключено**) и нажмите кнопку **OK** или **Применить**, чтобы применить обновление.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-278">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="2f0bd-279">Mac</span><span class="sxs-lookup"><span data-stu-id="2f0bd-279">Mac</span></span>](#tab/unix)

<span data-ttu-id="2f0bd-280">Надстройки Outlook для Mac, использующие функцию проверки при отправке, должны запускаться у всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-280">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="2f0bd-281">Однако если пользователям требуется запустить надстройку для выполнения стандартов соответствия требованиям, необходимо применить следующий параметр почтовых ящиков на компьютере каждого пользователя.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-281">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="2f0bd-282">Этот параметр или ключ совместим с CFPreference, то есть его можно установить, используя программное обеспечение для управления предприятием для Mac, например Jamf Pro.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-282">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

||<span data-ttu-id="2f0bd-283">Значение</span><span class="sxs-lookup"><span data-stu-id="2f0bd-283">Value</span></span>|
|:---|:---|
|<span data-ttu-id="2f0bd-284">**Домен**</span><span class="sxs-lookup"><span data-stu-id="2f0bd-284">**Domain**</span></span>|<span data-ttu-id="2f0bd-285">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-285">com.microsoft.outlook</span></span>|
|<span data-ttu-id="2f0bd-286">**Ключ**</span><span class="sxs-lookup"><span data-stu-id="2f0bd-286">**Key**</span></span>|<span data-ttu-id="2f0bd-287">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="2f0bd-287">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="2f0bd-288">**Тип данных**</span><span class="sxs-lookup"><span data-stu-id="2f0bd-288">**DataType**</span></span>|<span data-ttu-id="2f0bd-289">Логический</span><span class="sxs-lookup"><span data-stu-id="2f0bd-289">Boolean</span></span>|
|<span data-ttu-id="2f0bd-290">**Возможные значения**</span><span class="sxs-lookup"><span data-stu-id="2f0bd-290">**Possible values**</span></span>|<span data-ttu-id="2f0bd-291">false (по умолчанию)</span><span class="sxs-lookup"><span data-stu-id="2f0bd-291">false (default)</span></span><br><span data-ttu-id="2f0bd-292">true</span><span class="sxs-lookup"><span data-stu-id="2f0bd-292">true</span></span>|
|<span data-ttu-id="2f0bd-293">**Доступность**</span><span class="sxs-lookup"><span data-stu-id="2f0bd-293">**Availability**</span></span>|<span data-ttu-id="2f0bd-294">16.27</span><span class="sxs-lookup"><span data-stu-id="2f0bd-294">16.27</span></span>|
|<span data-ttu-id="2f0bd-295">**Примечания**</span><span class="sxs-lookup"><span data-stu-id="2f0bd-295">**Comments**</span></span>|<span data-ttu-id="2f0bd-296">Этот ключ создает политику onSendMailbox.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-296">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="2f0bd-297">Для чего служит параметр</span><span class="sxs-lookup"><span data-stu-id="2f0bd-297">What the setting does</span></span>

<span data-ttu-id="2f0bd-298">Для соответствия требованиям администраторам может потребоваться отключить возможность отправки сообщения или элементов собрания пользователями, пока не станет доступна для запуска последняя версия надстроек с функцией проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-298">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="2f0bd-299">Администраторы должны включить ключ **OnSendAddinsWaitForLoad**, чтобы все надстройки обновлялись из службы Exchange и были доступны для проверки того, что каждое сообщение или элемент собрания соответствует ожидаемым правилам и нормативным требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-299">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="2f0bd-300">Состояние ключа</span><span class="sxs-lookup"><span data-stu-id="2f0bd-300">Key's state</span></span>|<span data-ttu-id="2f0bd-301">Результат</span><span class="sxs-lookup"><span data-stu-id="2f0bd-301">Result</span></span>|
|---|---|
|<span data-ttu-id="2f0bd-302">false</span><span class="sxs-lookup"><span data-stu-id="2f0bd-302">false</span></span>|<span data-ttu-id="2f0bd-303">Загруженные в настоящее время манифесты надстройок (не обязательно последних версий) запускаются при отправке сообщений или элементов собраний.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-303">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="2f0bd-304">Это состояние по умолчанию/поведение.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-304">This is the default state/behavior.</span></span>|
|<span data-ttu-id="2f0bd-305">true</span><span class="sxs-lookup"><span data-stu-id="2f0bd-305">true</span></span>|<span data-ttu-id="2f0bd-306">После скачивания последних манифестов надстройки по отправке из Exchange надстройки запускаются на отправленных сообщениях или собраниях.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-306">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="2f0bd-307">В противном случае отправка блокируется, а кнопка **Отправка** отключена.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-307">Otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="2f0bd-308">Сценарии проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-308">On-send feature scenarios</span></span>

<span data-ttu-id="2f0bd-309">Ниже представлены поддерживаемые и неподдерживаемые сценарии для надстроек, использующих функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-309">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="2f0bd-310">В почтовом ящике пользователя включена функция проверки при отправке, но не установлено ни одной надстройки</span><span class="sxs-lookup"><span data-stu-id="2f0bd-310">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="2f0bd-311">В этом случае пользователь сможет отправлять сообщение или элементы собрания без запуска надстроек.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-311">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="2f0bd-312">В почтовом ящике пользователя включена функция проверки при отправке, а также установлены и включены надстройки, поддерживающие эту функцию</span><span class="sxs-lookup"><span data-stu-id="2f0bd-312">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="2f0bd-313">При отправке будут запускаться надстройки, которые разрешат или заблокируют отправку.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-313">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="2f0bd-314">Делегирование для почтовых ящиков, при котором у почтового ящика 1 есть разрешения на полный доступ к почтовому ящику 2</span><span class="sxs-lookup"><span data-stu-id="2f0bd-314">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="2f0bd-315">Веб-браузер (классическая версия Outlook)</span><span class="sxs-lookup"><span data-stu-id="2f0bd-315">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="2f0bd-316">Сценарий</span><span class="sxs-lookup"><span data-stu-id="2f0bd-316">Scenario</span></span>|<span data-ttu-id="2f0bd-317">Функция проверки при отправке для почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="2f0bd-317">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="2f0bd-318">Функция проверки при отправке для почтового ящика 2</span><span class="sxs-lookup"><span data-stu-id="2f0bd-318">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="2f0bd-319">Веб-сеанс Outlook (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="2f0bd-319">Outlook web session (classic)</span></span>|<span data-ttu-id="2f0bd-320">Результат</span><span class="sxs-lookup"><span data-stu-id="2f0bd-320">Result</span></span>|<span data-ttu-id="2f0bd-321">Поддержка</span><span class="sxs-lookup"><span data-stu-id="2f0bd-321">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="2f0bd-322">1</span><span class="sxs-lookup"><span data-stu-id="2f0bd-322">1</span></span>|<span data-ttu-id="2f0bd-323">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-323">Enabled</span></span>|<span data-ttu-id="2f0bd-324">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-324">Enabled</span></span>|<span data-ttu-id="2f0bd-325">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="2f0bd-325">New session</span></span>|<span data-ttu-id="2f0bd-326">Почтовый ящик 1 не может отправлять сообщение или элементы собраний из почтового ящика 2.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-326">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="2f0bd-p135">В настоящее время не поддерживается. В качестве обходного решения используйте сценарий 3.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="2f0bd-329">2</span><span class="sxs-lookup"><span data-stu-id="2f0bd-329">2</span></span>|<span data-ttu-id="2f0bd-330">Отключена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-330">Disabled</span></span>|<span data-ttu-id="2f0bd-331">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-331">Enabled</span></span>|<span data-ttu-id="2f0bd-332">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="2f0bd-332">New session</span></span>|<span data-ttu-id="2f0bd-333">Почтовый ящик 1 не может отправлять сообщение или элементы собраний из почтового ящика 2.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-333">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="2f0bd-p136">В настоящее время не поддерживается. В качестве обходного решения используйте сценарий 3.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-p136">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="2f0bd-336">3</span><span class="sxs-lookup"><span data-stu-id="2f0bd-336">3</span></span>|<span data-ttu-id="2f0bd-337">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-337">Enabled</span></span>|<span data-ttu-id="2f0bd-338">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-338">Enabled</span></span>|<span data-ttu-id="2f0bd-339">Тот же сеанс</span><span class="sxs-lookup"><span data-stu-id="2f0bd-339">Same session</span></span>|<span data-ttu-id="2f0bd-340">Проверка при отправке выполняется для почтового ящика 1, которому назначены надстройки, поддерживающие эту функцию.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-340">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="2f0bd-341">Поддерживается.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-341">Supported.</span></span>|
|<span data-ttu-id="2f0bd-342">4 </span><span class="sxs-lookup"><span data-stu-id="2f0bd-342">4</span></span>|<span data-ttu-id="2f0bd-343">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-343">Enabled</span></span>|<span data-ttu-id="2f0bd-344">Отключена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-344">Disabled</span></span>|<span data-ttu-id="2f0bd-345">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="2f0bd-345">New session</span></span>|<span data-ttu-id="2f0bd-346">Надстройки с функцией проверки при отправке не запускаются; отправка сообщения или элемента собрания.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-346">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="2f0bd-347">Поддерживается.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-347">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="2f0bd-348">Веб-браузер (современная версия Outlook), Windows, Mac</span><span class="sxs-lookup"><span data-stu-id="2f0bd-348">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="2f0bd-349">Чтобы внедрить функцию проверки при отправке, администраторы должны включить политику для обоих почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-349">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="2f0bd-350">Сведения о том, как поддерживать делегированный доступ в надстройке, см. в статье [Включение сценариев делегированного доступа в надстройке Outlook](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-350">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="2f0bd-351">Группа 1 — это современный почтовый ящик группы, а почтовый ящик пользователя 1 является участником группы 1</span><span class="sxs-lookup"><span data-stu-id="2f0bd-351">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="2f0bd-352">Сценарий</span><span class="sxs-lookup"><span data-stu-id="2f0bd-352">Scenario</span></span>|<span data-ttu-id="2f0bd-353">Политика проверки при отправке для почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="2f0bd-353">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="2f0bd-354">Включены ли надстройки, поддерживающие проверку сообщений при отправке?</span><span class="sxs-lookup"><span data-stu-id="2f0bd-354">On-send add-ins enabled?</span></span>|<span data-ttu-id="2f0bd-355">Действие почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="2f0bd-355">Mailbox 1 action</span></span>|<span data-ttu-id="2f0bd-356">Результат</span><span class="sxs-lookup"><span data-stu-id="2f0bd-356">Result</span></span>|<span data-ttu-id="2f0bd-357">Поддержка</span><span class="sxs-lookup"><span data-stu-id="2f0bd-357">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="2f0bd-358">1</span><span class="sxs-lookup"><span data-stu-id="2f0bd-358">1</span></span>|<span data-ttu-id="2f0bd-359">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-359">Enabled</span></span>|<span data-ttu-id="2f0bd-360">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-360">Yes</span></span>|<span data-ttu-id="2f0bd-361">Почтовый ящик 1 создает новое сообщение или собрание для группы 1.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-361">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="2f0bd-362">В случае отправки запускаются надстройки, поддерживающие проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-362">On-send add-ins run during send.</span></span>|<span data-ttu-id="2f0bd-363">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-363">Yes</span></span>|
|<span data-ttu-id="2f0bd-364">2</span><span class="sxs-lookup"><span data-stu-id="2f0bd-364">2</span></span>|<span data-ttu-id="2f0bd-365">Включена</span><span class="sxs-lookup"><span data-stu-id="2f0bd-365">Enabled</span></span>|<span data-ttu-id="2f0bd-366">Да</span><span class="sxs-lookup"><span data-stu-id="2f0bd-366">Yes</span></span>|<span data-ttu-id="2f0bd-367">Почтовый ящик 1 создает новое сообщение или собрание для группы 1 в окне этой группы в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-367">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="2f0bd-368">В случае отправки не запускаются надстройки, поддерживающие проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-368">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="2f0bd-369">В настоящее время не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-369">Not currently supported.</span></span> <span data-ttu-id="2f0bd-370">В качестве обходного решения используйте сценарий 1.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-370">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="2f0bd-371">Включен почтовый ящик пользователя с функцией или политикой проверки при отправке, установлены и включены надстройки, поддерживающие эту функцию, а также включен автономный режим</span><span class="sxs-lookup"><span data-stu-id="2f0bd-371">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="2f0bd-372">Надстройки, поддерживающие проверку при отправке, запускаются в соответствии с сетевым состоянием пользователя, внутреннего сервера надстройки и Exchange.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-372">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="2f0bd-373">Состояние пользователя</span><span class="sxs-lookup"><span data-stu-id="2f0bd-373">User's state</span></span>

<span data-ttu-id="2f0bd-374">Надстройки, поддерживающие проверку сообщений при отправке, будут запускаться при отправке, если пользователь в сети.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-374">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="2f0bd-375">В автономном режиме такие надстройки не будут запускаться при отправке, а сообщение или элемент собрания не будет отправлен.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-375">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="2f0bd-376">Состояние внутреннего сервера надстройки</span><span class="sxs-lookup"><span data-stu-id="2f0bd-376">Add-in backend's state</span></span>

<span data-ttu-id="2f0bd-377">Надстройка, поддерживающая проверку при отправке, будет запускаться, если ее внутренний сервер подключен к сети и доступен.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-377">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="2f0bd-378">Если внутренний сервер находится в автономном режиме, отправка отключена.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-378">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="2f0bd-379">Состояние Exchange</span><span class="sxs-lookup"><span data-stu-id="2f0bd-379">Exchange's state</span></span>

<span data-ttu-id="2f0bd-380">Надстройки, поддерживающие проверку сообщений при отправке, будут запускаться при отправке, если сервер Exchange подключен к сети и доступен.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-380">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="2f0bd-381">Если надстройке с функцией проверки при отправке недоступна служба Exchange и включена соответствующая политика или командлет, отправка отключена.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-381">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="2f0bd-382">На компьютерах Mac в любом автономном состоянии кнопка **Отправить** (или **Отправить обновление** для существующих собраний) отключена, и отображается уведомление, что в организации не разрешено отправлять сообщения, если пользователь не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-382">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="2f0bd-383">Пользователь может изменять элемент во время работы надстройки при отправке</span><span class="sxs-lookup"><span data-stu-id="2f0bd-383">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="2f0bd-384">В то время как надстройки при отправке обрабатывают элемент, пользователь может изменить элемент, добавив, например, неправильный текст или вложения.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-384">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="2f0bd-385">Если вы хотите запретить пользователю изменять элемент во время обработки надстройки при отправке, вы можете реализовать обходное решение с помощью диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-385">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="2f0bd-386">Это обходное решение можно использовать в Outlook в Интернете (классический), Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-386">This workaround can be used in Outlook on the web (classic), Windows, and Mac.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f0bd-387">Современный outlook в Интернете. Чтобы предотвратить редактирование элемента во время обработки надстройки при отправке, необходимо установить флаг *OnSendAddinsEnabled,* как описано в надстройке `true` [Install Outlook,](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) которая используется в разделе отправки в начале этой статьи.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-387">Modern Outlook on the web: To prevent the user from editing the item while your add-in is processing on send, you should set the *OnSendAddinsEnabled* flag to `true` as described in the [Install Outlook add-ins that use on-send](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) section earlier in this article.</span></span>

<span data-ttu-id="2f0bd-388">В обработнике отправки:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-388">In your on-send handler:</span></span>

1. <span data-ttu-id="2f0bd-389">Вызов [displayDialogAsync,](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) чтобы открыть диалоговое окно, чтобы щелчки мыши и клавиши были отключены.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-389">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="2f0bd-390">Чтобы получить это поведение в классическом Outlook в Интернете, необходимо настроить свойство [displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) в `true` `options` параметре `displayDialogAsync` вызова.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-390">To get this behavior in classic Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="2f0bd-391">Реализация обработки элемента.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-391">Implement processing of the item.</span></span>
1. <span data-ttu-id="2f0bd-392">Закройте диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-392">Close the dialog.</span></span> <span data-ttu-id="2f0bd-393">Кроме того, обработать, что произойдет, если пользователь закрывает диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-393">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="2f0bd-394">Примеры кода</span><span class="sxs-lookup"><span data-stu-id="2f0bd-394">Code examples</span></span>

<span data-ttu-id="2f0bd-395">В приведенных ниже примерах кода показано, как создать простую надстройку, поддерживающую проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-395">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="2f0bd-396">Скачать код, на котором основаны эти примеры, можно на странице [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-396">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="2f0bd-397">Если вы используете диалоговое окно с событием отправки, не забудьте закрыть диалоговое окно перед завершением события.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-397">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="2f0bd-398">Манифест, переопределение версии и событие</span><span class="sxs-lookup"><span data-stu-id="2f0bd-398">Manifest, version override, and event</span></span>

<span data-ttu-id="2f0bd-399">Пример кода [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) включает два манифеста:</span><span class="sxs-lookup"><span data-stu-id="2f0bd-399">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="2f0bd-400">`Contoso Message Body Checker.xml` показывает, как проверить текст сообщения на наличие запрещенных слов или конфиденциальной информации при отправке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-400">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="2f0bd-401">`Contoso Subject and CC Checker.xml` показывает, как при отправке добавить получателя в строку "Копия" и проверить, включает ли сообщение строку темы.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-401">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="2f0bd-402">В файле манифеста `Contoso Message Body Checker.xml` указываются файл и имя функции, которую следует вызывать при возникновении события `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-402">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="2f0bd-403">Операция выполняется синхронно.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-403">The operation runs synchronously.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case, the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

> [!IMPORTANT]
> <span data-ttu-id="2f0bd-404">Если вы используете Visual Studio 2019 для разработки надстройки при отправке, вы можете получить предупреждение о проверке, как следующее: "Это недействительный xsi:type http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events '". Чтобы обойти это, вам потребуется более новая версия MailAppVersionOverridesV1_1.xsd, которая была предоставлена в качестве gist GitHub в блоге об этом [предупреждении](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-404">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="2f0bd-405">Для файла манифеста `Contoso Subject and CC Checker.xml` в приведенном ниже примере показаны файл и имя функции, вызываемой при возникновении события отправки.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-405">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

<br/>

<span data-ttu-id="2f0bd-406">Для API проверки при отправке требуется узел `VersionOverrides v1_1`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-406">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="2f0bd-407">Ниже показано, как добавить узел `VersionOverrides` в манифест.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-407">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="2f0bd-408">Дополнительные сведения см. в указанных ниже статьях.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-408">For more information, see the following:</span></span>
> - [<span data-ttu-id="2f0bd-409">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-409">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="2f0bd-410">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="2f0bd-410">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="2f0bd-411">Объекты `Event` и `item`, методы `body.getAsync` и `body.setAsync`</span><span class="sxs-lookup"><span data-stu-id="2f0bd-411">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="2f0bd-412">Чтобы получить доступ к выбранному в данный момент сообщению или элементу собрания (в этом примере — к новому сообщению), используйте пространство имен `Office.context.mailbox.item`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-412">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="2f0bd-413">Функция проверки при отправке автоматически передает событие `ItemSend` функции, указанной в манифесте (в данном случае это функция `validateBody`).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-413">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

```js
var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
```

<span data-ttu-id="2f0bd-414">Функция `validateBody` возвращает текущий текст в заданном формате (HTML) и передает нужный объект события `ItemSend` в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-414">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="2f0bd-415">Помимо метода `getAsync`, объект `Body` также предоставляет метод `setAsync`, с помощью которого вы можете заменить текст сообщения на указанный.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-415">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="2f0bd-416">Дополнительные сведения см. в статьях [Объект Event](/javascript/api/office/office.addincommands.event) и [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-416">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="2f0bd-417">Объект `NotificationMessages` и метод `event.completed`</span><span class="sxs-lookup"><span data-stu-id="2f0bd-417">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="2f0bd-418">Функция `checkBodyOnlyOnSendCallBack` использует регулярное выражение, чтобы определить, содержит ли текст сообщения слова, подлежащие блокировке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-418">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="2f0bd-419">Если она обнаруживает слово, совпадающие с каким-либо элементом из массива запрещенных слов, отправка сообщения блокируется, а отправитель получает уведомление на панели информации.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-419">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="2f0bd-420">Для этого в ней используется свойство `notificationMessages` объекта `Item` для возврата объекта `NotificationMessages`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-420">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="2f0bd-421">После этого она добавляет уведомление к элементу, вызывая метод `addAsync`, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-421">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

```js
// Determine whether the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allow sending.
// <param name="asyncResult">ItemSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    var wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
}
```

<span data-ttu-id="2f0bd-422">Ниже перечислены параметры метода `addAsync`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-422">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="2f0bd-423">`NoSend`. Строка, представляющая собой заданный разработчиком ключ для ссылки на сообщение уведомления.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-423">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="2f0bd-424">С его помощью вы сможете изменить это сообщение позже.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-424">You can use it to modify this message later.</span></span> <span data-ttu-id="2f0bd-425">Клавиша не может быть длиннее 32 символов.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-425">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="2f0bd-426">`type`. Одно из свойств параметра объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-426">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="2f0bd-427">Представляет тип сообщения. Типы соответствуют значениям перечисления [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-427">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="2f0bd-428">Допустимые значения: индикатор хода выполнения, информационное сообщение и сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-428">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="2f0bd-429">В этом примере в свойстве `type` указано сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-429">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="2f0bd-430">`message`. Одно из свойств параметра объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-430">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="2f0bd-431">В этом примере `message` — это текст сообщения уведомления.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-431">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="2f0bd-432">Чтобы сообщить о завершении надстройкой обработки события `ItemSend`, активированного операцией отправки, вызовите метод `event.completed({allowEvent:Boolean})`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-432">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="2f0bd-433">Свойство `allowEvent` является логическим.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-433">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="2f0bd-434">Если задано значение `true`, отправка разрешается.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-434">If set to `true`, send is allowed.</span></span> <span data-ttu-id="2f0bd-435">Если задано значение `false`, отправка письма блокируется.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-435">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="2f0bd-436">Дополнительные сведения см. в статьях [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [completed](/javascript/api/office/office.addincommands.event).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-436">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="2f0bd-437">Методы `replaceAsync`, `removeAsync` и `getAllAsync`</span><span class="sxs-lookup"><span data-stu-id="2f0bd-437">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="2f0bd-438">Помимо метода `addAsync`, объект `NotificationMessages` также включает методы `replaceAsync`, `removeAsync` и `getAllAsync`.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-438">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="2f0bd-439">Эти методы не используются в данном примере кода.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-439">These methods are not used in this code sample.</span></span>  <span data-ttu-id="2f0bd-440">Дополнительные сведения см. в статье [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span><span class="sxs-lookup"><span data-stu-id="2f0bd-440">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="2f0bd-441">Код проверки строк "Тема" и "Копия"</span><span class="sxs-lookup"><span data-stu-id="2f0bd-441">Subject and CC checker code</span></span>

<span data-ttu-id="2f0bd-442">В приведенном ниже примере кода показано, как при отправке сообщения добавить получателя в строку "Копия" и проверить, включает ли сообщение тему.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-442">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="2f0bd-443">В этом примере функция проверки при отправке используется, чтобы разрешить или запретить отправку сообщения.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-443">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

```js
// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

// Determine whether the subject should be changed. If it is already changed, allow send. Otherwise change it.
// <param name="event">ItemSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Determine whether a string is blank, null, or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
        });
}

// Add a CC to the email. In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">ItemSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Determine whether the subject should be changed. If it is already changed, allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">ItemSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        });
}
```

<span data-ttu-id="2f0bd-p156">Дополнительные сведения о том, как при отправке сообщения добавить получателя в строку "Копия" и проверить, указана ли тема сообщения, а также просмотреть доступные API, см. в [примере Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). Код сопровождается подробными комментариями.</span><span class="sxs-lookup"><span data-stu-id="2f0bd-p156">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="2f0bd-446">См. также</span><span class="sxs-lookup"><span data-stu-id="2f0bd-446">See also</span></span>

- [<span data-ttu-id="2f0bd-447">Обзор архитектуры и функций надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="2f0bd-447">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="2f0bd-448">Надстройка Outlook "Демонстрация команд надстройки"</span><span class="sxs-lookup"><span data-stu-id="2f0bd-448">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)