---
title: Функция проверки при отправке для надстроек Outlook
description: Позволяет надстройке настраивать те или иные параметры при отправке, а также обрабатывать элемент и запрещать пользователям выполнять определенные действия.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 09bc44e78f202474757317b4b07bce50d6235aa1
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836902"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="eeea4-103">Функция проверки при отправке для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="eeea4-p101">Функция проверки при отправке позволяет надстройкам Outlook настраивать те или иные параметры при отправке, а также обрабатывать сообщение или элемент собрания и запрещать пользователям выполнять определенные действия. Например, с помощью функции проверки при отправке можно сделать следующее:</span><span class="sxs-lookup"><span data-stu-id="eeea4-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="eeea4-106">запретить пользователю отправлять конфиденциальную информацию или оставлять строку темы пустой;</span><span class="sxs-lookup"><span data-stu-id="eeea4-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="eeea4-107">добавить определенного получателя в строку "Копия" в сообщениях или в строку "Необязательные получатели" в собраниях.</span><span class="sxs-lookup"><span data-stu-id="eeea4-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="eeea4-108">Функция проверки при отправке не использует пользовательский интерфейс и активируется событием типа `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="eeea4-109">Сведения об ограничениях, связанных с функцией проверки при отправке, см. в разделе [Ограничения](#limitations) далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="eeea4-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="eeea4-110">Поддерживаемые клиенты и платформы</span><span class="sxs-lookup"><span data-stu-id="eeea4-110">Supported clients and platforms</span></span>

<span data-ttu-id="eeea4-111">В следующей таблице показаны поддерживаемые комбинации клиент-сервер для функции отправки.</span><span class="sxs-lookup"><span data-stu-id="eeea4-111">The following table shows supported client-server combinations for the on-send feature.</span></span> <span data-ttu-id="eeea4-112">Исключенные комбинации не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="eeea4-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="eeea4-113">Client</span><span class="sxs-lookup"><span data-stu-id="eeea4-113">Client</span></span> | <span data-ttu-id="eeea4-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="eeea4-114">Exchange Online</span></span> | <span data-ttu-id="eeea4-115">Локальное exchange 2016</span><span class="sxs-lookup"><span data-stu-id="eeea4-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="eeea4-116">(Накопительное обновление 6 или более позднее)</span><span class="sxs-lookup"><span data-stu-id="eeea4-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="eeea4-117">Локальное exchange 2019</span><span class="sxs-lookup"><span data-stu-id="eeea4-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="eeea4-118">(Накопительное обновление 1 или более позднее)</span><span class="sxs-lookup"><span data-stu-id="eeea4-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="eeea4-119">Windows:</span><span class="sxs-lookup"><span data-stu-id="eeea4-119">Windows:</span></span><br><span data-ttu-id="eeea4-120">версия 1910 (сборка 12130.20272) или более поздней версии</span><span class="sxs-lookup"><span data-stu-id="eeea4-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="eeea4-121">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-121">Yes</span></span>|<span data-ttu-id="eeea4-122">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-122">Yes</span></span>|<span data-ttu-id="eeea4-123">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-123">Yes</span></span>|
|<span data-ttu-id="eeea4-124">Mac:</span><span class="sxs-lookup"><span data-stu-id="eeea4-124">Mac:</span></span><br><span data-ttu-id="eeea4-125">сборка 16.30 или более поздней</span><span class="sxs-lookup"><span data-stu-id="eeea4-125">build 16.30 or later</span></span>|<span data-ttu-id="eeea4-126">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-126">Yes</span></span>|<span data-ttu-id="eeea4-127">Нет</span><span class="sxs-lookup"><span data-stu-id="eeea4-127">No</span></span>|<span data-ttu-id="eeea4-128">Нет</span><span class="sxs-lookup"><span data-stu-id="eeea4-128">No</span></span>|
|<span data-ttu-id="eeea4-129">Веб-браузер:</span><span class="sxs-lookup"><span data-stu-id="eeea4-129">Web browser:</span></span><br><span data-ttu-id="eeea4-130">современный пользовательский интерфейс Outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-130">modern Outlook UI</span></span>|<span data-ttu-id="eeea4-131">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-131">Yes</span></span>|<span data-ttu-id="eeea4-132">Неприменимо</span><span class="sxs-lookup"><span data-stu-id="eeea4-132">Not applicable</span></span>|<span data-ttu-id="eeea4-133">Неприменимо</span><span class="sxs-lookup"><span data-stu-id="eeea4-133">Not applicable</span></span>|
|<span data-ttu-id="eeea4-134">Веб-браузер:</span><span class="sxs-lookup"><span data-stu-id="eeea4-134">Web browser:</span></span><br><span data-ttu-id="eeea4-135">классический пользовательский интерфейс Outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-135">classic Outlook UI</span></span>|<span data-ttu-id="eeea4-136">Неприменимо</span><span class="sxs-lookup"><span data-stu-id="eeea4-136">Not applicable</span></span>|<span data-ttu-id="eeea4-137">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-137">Yes</span></span>|<span data-ttu-id="eeea4-138">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-138">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="eeea4-139">Функция отправки была официально выпущена в наборе требований 1.8 (см. сведения о текущем сервере [и клиентской](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) поддержке).</span><span class="sxs-lookup"><span data-stu-id="eeea4-139">The on-send feature was officially released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span> <span data-ttu-id="eeea4-140">Однако обратите внимание, что матрица поддержки функции является суперсетью набора требований.</span><span class="sxs-lookup"><span data-stu-id="eeea4-140">However, note that the feature's support matrix is a superset of the requirement set's.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eeea4-141">Надстройки, которые используют функцию отправки, не допускаются в [AppSource.](https://appsource.microsoft.com)</span><span class="sxs-lookup"><span data-stu-id="eeea4-141">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="eeea4-142">Как работает функция проверки при отправке?</span><span class="sxs-lookup"><span data-stu-id="eeea4-142">How does the on-send feature work?</span></span>

<span data-ttu-id="eeea4-143">С помощью функции проверки при отправке вы можете создать надстройку Outlook, задействующую синхронное событие `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-143">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="eeea4-144">Это событие возникает, когда пользователь нажимает кнопку **Отправить** (или **Отправить обновление** для существующих собраний). С его помощью можно блокировать отправку элемента, не прошедшего проверку.</span><span class="sxs-lookup"><span data-stu-id="eeea4-144">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="eeea4-145">Например, когда пользователь вызывает событие отправки сообщения, надстройка Outlook благодаря функции проверки при отправке может следующее:</span><span class="sxs-lookup"><span data-stu-id="eeea4-145">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="eeea4-146">прочитать и проверить содержимое письма;</span><span class="sxs-lookup"><span data-stu-id="eeea4-146">Read and validate the email message contents</span></span>
- <span data-ttu-id="eeea4-147">проверить, содержит ли сообщение строку темы;</span><span class="sxs-lookup"><span data-stu-id="eeea4-147">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="eeea4-148">задать заранее определенного получателя.</span><span class="sxs-lookup"><span data-stu-id="eeea4-148">Set a predetermined recipient</span></span>

<span data-ttu-id="eeea4-149">Проверка проводится на стороне клиента в Outlook, когда событие отправки запускается, и надстройка имеет до 5 минут, прежде чем он раз. Если проверка не удается, отправка элемента блокируется, а сообщение об ошибке отображается в информационной панели, которая побуждает пользователя к действию.</span><span class="sxs-lookup"><span data-stu-id="eeea4-149">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

> [!NOTE]
> <span data-ttu-id="eeea4-150">В Outlook в Интернете, когда функция отправки запускается в сообщении, сочиняемом в вкладке браузера Outlook, элемент выскочил в собственное окно браузера или вкладку, чтобы завершить проверку и другую обработку.</span><span class="sxs-lookup"><span data-stu-id="eeea4-150">In Outlook on the web, when the on-send feature is triggered in a message being composed within the Outlook browser tab, the item is popped out to its own browser window or tab in order to complete validation and other processing.</span></span>

<span data-ttu-id="eeea4-151">На приведенном ниже снимке экрана показана панель информации, в которой пользователю предлагается добавить тему.</span><span class="sxs-lookup"><span data-stu-id="eeea4-151">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![Снимок экрана с сообщением об ошибке, в котором пользователю предлагается ввести отсутствующую строку темы](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="eeea4-153">На приведенном ниже снимке экрана показана панель информации, уведомляющая отправителя о том, что обнаружены слова, подлежащие блокировке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-153">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![Снимок экрана с сообщением об ошибке, которое сообщает пользователю о том, что обнаружены слова, подлежащие блокировке](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="eeea4-155">Ограничения</span><span class="sxs-lookup"><span data-stu-id="eeea4-155">Limitations</span></span>

<span data-ttu-id="eeea4-156">В настоящее время на функцию проверки при отправке действуют перечисленные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="eeea4-156">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="eeea4-157">**Функция append-on-send** &ndash; При вызове [тела. В обработнике отправки appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="eeea4-157">**Append-on-send** feature &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="eeea4-158">**AppSource**. Надстройки Outlook, в которых используется функция проверки при отправке, невозможно публиковать в [AppSource](https://appsource.microsoft.com), так как они не проходят проверку AppSource.</span><span class="sxs-lookup"><span data-stu-id="eeea4-158">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="eeea4-159">Надстройки, использующие функцию проверки при отправке, должны разворачиваться администраторами.</span><span class="sxs-lookup"><span data-stu-id="eeea4-159">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="eeea4-160">**Манифест**. Для каждой надстройки поддерживается только одно событие `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-160">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="eeea4-161">Если манифест содержит несколько событий `ItemSend`, он не пройдет проверку.</span><span class="sxs-lookup"><span data-stu-id="eeea4-161">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="eeea4-p107">**Производительность**. &ndash;Многочисленные случаи приема-передачи пакетов на веб-сервере, где размещается надстройка, могут повлиять на ее производительность. Учитывайте влияние на производительность при создании надстроек, требующих выполнения нескольких операций с сообщениями или собраниями.</span><span class="sxs-lookup"><span data-stu-id="eeea4-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="eeea4-164">**Отправить позже** (только для Mac). Если имеются надстройки, использующие проверку при отправке, функция **Отправить позже** будет недоступна.</span><span class="sxs-lookup"><span data-stu-id="eeea4-164">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

<span data-ttu-id="eeea4-165">Кроме того, не рекомендуется вызывать обработитель событий при отправке, так как закрытие элемента должно происходить автоматически после `item.close()` завершения события.</span><span class="sxs-lookup"><span data-stu-id="eeea4-165">Also, it's not recommended that you call `item.close()` in the on-send event handler as closing the item should happen automatically after the event is completed.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="eeea4-166">Ограничения на типы и режимы почтовых ящиков</span><span class="sxs-lookup"><span data-stu-id="eeea4-166">Mailbox type/mode limitations</span></span>

<span data-ttu-id="eeea4-167">Функция проверки при отправке поддерживается только для почтовых ящиков пользователей в Outlook в Интернете, Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="eeea4-167">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="eeea4-168">В настоящее время она не поддерживается для указанных ниже типов и режимов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-168">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="eeea4-169">Общие почтовые ящики\*</span><span class="sxs-lookup"><span data-stu-id="eeea4-169">Shared mailboxes\*</span></span>
- <span data-ttu-id="eeea4-170">почтовые ящики групп;</span><span class="sxs-lookup"><span data-stu-id="eeea4-170">Group mailboxes</span></span>
- <span data-ttu-id="eeea4-171">почтовые ящики в автономном режиме.</span><span class="sxs-lookup"><span data-stu-id="eeea4-171">Offline mode</span></span>

<span data-ttu-id="eeea4-172">Outlook не разрешает отправку, если для этих сценариев почтовых ящиков включена функция проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-172">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="eeea4-173">Однако если пользователь отвечает на сообщение в почтовом ящике группы, надстройка, поддерживающая проверку сообщений при отправке, не запускается, а сообщение отправляется.</span><span class="sxs-lookup"><span data-stu-id="eeea4-173">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eeea4-174">\*Функции отправки должны работать на общих почтовых ящиках или папках, если надстройка также реализует поддержку сценариев доступа [к делегатам.](delegate-access.md)</span><span class="sxs-lookup"><span data-stu-id="eeea4-174">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="eeea4-175">Несколько надстроек, поддерживающих проверку сообщений при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-175">Multiple on-send add-ins</span></span>

<span data-ttu-id="eeea4-176">Если установлено несколько надстроек, поддерживающих проверку сообщений при отправке, они будут запускаться в том порядке, в котором были получены из API `getAppManifestCall` или `getExtensibilityContext`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-176">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="eeea4-177">Если первая надстройка разрешает отправку сообщения, то вторая может внести изменения, на которые первая отреагировала бы запретом отправки.</span><span class="sxs-lookup"><span data-stu-id="eeea4-177">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="eeea4-178">Однако первая надстройка не будет запускаться снова, если все установленные надстройки разрешат отправку.</span><span class="sxs-lookup"><span data-stu-id="eeea4-178">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="eeea4-179">Допустим, надстройка 1 и надстройка 2 используют функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-179">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="eeea4-180">Сначала устанавливается надстройка 1, а затем — надстройка 2.</span><span class="sxs-lookup"><span data-stu-id="eeea4-180">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="eeea4-181">Надстройка 1 находит в сообщении слово Fabrikam, что является условием для разрешения отправки.</span><span class="sxs-lookup"><span data-stu-id="eeea4-181">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="eeea4-182">Однако надстройка 2 удаляет все вхождения слова Fabrikam.</span><span class="sxs-lookup"><span data-stu-id="eeea4-182">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="eeea4-183">Сообщение будет отправлено без единого слова Fabrikam (в связи с порядком установки надстроек 1 и 2).</span><span class="sxs-lookup"><span data-stu-id="eeea4-183">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="eeea4-184">Развертывание надстроек Outlook, использующих функцию проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-184">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="eeea4-185">Рекомендуем, чтобы развертывание надстроек Outlook, использующих функцию проверки при отправке, выполняли администраторы.</span><span class="sxs-lookup"><span data-stu-id="eeea4-185">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="eeea4-186">Администратор должен убедиться, что такая надстройка:</span><span class="sxs-lookup"><span data-stu-id="eeea4-186">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="eeea4-187">всегда присутствует при открытии создаваемого элемента (для электронной почты: создание сообщений, ответ и пересылка);</span><span class="sxs-lookup"><span data-stu-id="eeea4-187">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="eeea4-188">не может быть закрыта или отключена пользователем.</span><span class="sxs-lookup"><span data-stu-id="eeea4-188">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="eeea4-189">Установка надстроек Outlook, использующих функцию проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-189">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="eeea4-190">Чтобы использовать функцию проверки при отправке в Outlook, надстройки должны быть настроены для типов событий отправки.</span><span class="sxs-lookup"><span data-stu-id="eeea4-190">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="eeea4-191">Выберите платформу, которую нужно настроить.</span><span class="sxs-lookup"><span data-stu-id="eeea4-191">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="eeea4-192">Веб-браузер — классическая версия Outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-192">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="eeea4-193">Надстройки для Outlook в Интернете (классическая версия), использующие функцию проверки при отправке, будут запускаться у пользователей, которым назначена политика почтовых ящиков Outlook в Интернете, для флага *OnSendAddinsEnabled* которой задано значение **true**.</span><span class="sxs-lookup"><span data-stu-id="eeea4-193">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="eeea4-194">Чтобы установить новую надстройку, выполните приведенные ниже командлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="eeea4-194">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="eeea4-195">Сведения о том, как подключиться к Exchange Online с помощью удаленного сеанса PowerShell, см. в статье [Подключение к Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="eeea4-195">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="eeea4-196">Включение функции проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-196">Enable the on-send feature</span></span>

<span data-ttu-id="eeea4-197">По умолчанию функция проверки при отправке отключена.</span><span class="sxs-lookup"><span data-stu-id="eeea4-197">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="eeea4-198">Администраторы могут включать эту функцию с помощью командлетов Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="eeea4-198">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="eeea4-199">Чтобы для всех пользователей включить надстройки, поддерживающие проверку сообщений при отправке, сделайте следующее:</span><span class="sxs-lookup"><span data-stu-id="eeea4-199">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="eeea4-200">Создайте политику почтовых ящиков Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="eeea4-200">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="eeea4-201">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-201">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="eeea4-202">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-202">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="eeea4-203">Включите функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-203">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="eeea4-204">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="eeea4-204">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="eeea4-205">Включение функции проверки при отправке для группы пользователей</span><span class="sxs-lookup"><span data-stu-id="eeea4-205">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="eeea4-206">Чтобы включить функцию проверки при отправке для определенной группы пользователей, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="eeea4-206">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="eeea4-207">В этом примере администратор включает функцию проверки при отправке для надстроек Outlook в Интернете только в среде финансового отдела (Finance).</span><span class="sxs-lookup"><span data-stu-id="eeea4-207">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="eeea4-208">Создайте политику почтовых ящиков Outlook в Интернете для группы.</span><span class="sxs-lookup"><span data-stu-id="eeea4-208">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="eeea4-209">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков (дополнительные сведения см. в разделе [Ограничения на типы почтовых ящиков](#multiple-on-send-add-ins) выше в этой статье).</span><span class="sxs-lookup"><span data-stu-id="eeea4-209">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="eeea4-210">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-210">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="eeea4-211">Включите функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-211">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="eeea4-212">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="eeea4-212">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="eeea4-213">Дождитесь вступления политики в силу (это может занять до 60 минут) или перезапустите службы IIS.</span><span class="sxs-lookup"><span data-stu-id="eeea4-213">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="eeea4-214">Когда политика вступит в силу, для группы будет включена функция проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-214">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="eeea4-215">Отключение функции проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-215">Disable the on-send feature</span></span>

<span data-ttu-id="eeea4-216">Чтобы отключить функцию проверки при отправке для пользователя или назначить политику почтовых ящиков Outlook в Интернете, в которой не включен соответствующий флаг, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="eeea4-216">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="eeea4-217">В этом примере используется политика почтовых ящиков *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="eeea4-217">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="eeea4-218">Дополнительные сведения о том, как настроить существующие политики почтовых ящиков Outlook в Интернете с помощью командлета **Set-OwaMailboxPolicy**, см. в статье [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="eeea4-218">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="eeea4-219">Чтобы отключить функцию проверки при отправке для всех пользователей, которым назначена определенная политика почтовых ящиков Outlook в Интернете, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="eeea4-219">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="eeea4-220">Веб-браузер — современная версия Outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-220">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="eeea4-221">Надстройки для Outlook в Интернете (современная версия), использующие функцию проверки при отправке, должны запускаться для всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="eeea4-221">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="eeea4-222">Однако если пользователям необходимо запускать надстройки по отправке в соответствии со стандартами соответствия требованиям, политика почтовых ящиков должна иметь флаг *OnSendAddinsEnabled,* чтобы редактирование элемента не было разрешено во время обработки надстройок при `true` отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-222">However, if users are required to run on-send add-ins to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to `true` so that editing the item is not allowed while the add-ins are processing on send.</span></span>

<span data-ttu-id="eeea4-223">Чтобы установить новую надстройку, выполните приведенные ниже командлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="eeea4-223">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="eeea4-224">Сведения о том, как подключиться к Exchange Online с помощью удаленного сеанса PowerShell, см. в статье [Подключение к Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="eeea4-224">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-flag"></a><span data-ttu-id="eeea4-225">Включить флаг отправки</span><span class="sxs-lookup"><span data-stu-id="eeea4-225">Enable the on-send flag</span></span>

<span data-ttu-id="eeea4-226">Администраторы могут обеспечивать соответствие требованиям при отправке, запуская комдлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="eeea4-226">Administrators can enforce on-send compliance by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="eeea4-227">Чтобы отослать редактирование во время обработки надстройок при отправке, для всех пользователей необходимо отослать:</span><span class="sxs-lookup"><span data-stu-id="eeea4-227">For all users, to disallow editing while on-send add-ins are processing:</span></span>

1. <span data-ttu-id="eeea4-228">Создайте политику почтовых ящиков Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="eeea4-228">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="eeea4-229">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-229">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="eeea4-230">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-230">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="eeea4-231">Обеспечение соответствия требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-231">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="eeea4-232">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="eeea4-232">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a><span data-ttu-id="eeea4-233">Включив флаг отправки для группы пользователей</span><span class="sxs-lookup"><span data-stu-id="eeea4-233">Turn on the on-send flag for a group of users</span></span>

<span data-ttu-id="eeea4-234">Чтобы обеспечить соблюдение при отправке для определенной группы пользователей, выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="eeea4-234">To enforce on-send compliance for a specific group of users, the steps are as follows.</span></span> <span data-ttu-id="eeea4-235">В этом примере администратор включает политику проверки при отправке для надстроек Outlook в Интернете только в среде финансового отдела (Finance).</span><span class="sxs-lookup"><span data-stu-id="eeea4-235">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="eeea4-236">Создайте политику почтовых ящиков Outlook в Интернете для группы.</span><span class="sxs-lookup"><span data-stu-id="eeea4-236">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="eeea4-237">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков (дополнительные сведения см. в разделе [Ограничения на типы почтовых ящиков](#multiple-on-send-add-ins) выше в этой статье).</span><span class="sxs-lookup"><span data-stu-id="eeea4-237">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="eeea4-238">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-238">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="eeea4-239">Обеспечение соответствия требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-239">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="eeea4-240">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="eeea4-240">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="eeea4-241">Дождитесь вступления политики в силу (это может занять до 60 минут) или перезапустите службы IIS.</span><span class="sxs-lookup"><span data-stu-id="eeea4-241">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="eeea4-242">Когда политика вступает в силу, для группы будет применено соответствие требованиям по отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-242">When the policy takes effect, on-send compliance will be enforced for the group.</span></span>

#### <a name="turn-off-the-on-send-flag"></a><span data-ttu-id="eeea4-243">Выключите флаг отправки</span><span class="sxs-lookup"><span data-stu-id="eeea4-243">Turn off the on-send flag</span></span>

<span data-ttu-id="eeea4-244">Чтобы отключить для пользователя правоприменительных требований к выполнению при отправке, назначьте Outlook в политике веб-почтовых ящиков, в которую не включен флаг, запуская следующие cmdlets.</span><span class="sxs-lookup"><span data-stu-id="eeea4-244">To turn off on-send compliance enforcement for a user, assign an Outlook on the web mailbox policy that does not have the flag enabled by running the following cmdlets.</span></span> <span data-ttu-id="eeea4-245">В этом примере используется политика почтовых ящиков *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="eeea4-245">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="eeea4-246">Дополнительные сведения о том, как настроить существующие политики почтовых ящиков Outlook в Интернете с помощью командлета **Set-OwaMailboxPolicy**, см. в статье [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="eeea4-246">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="eeea4-247">Чтобы отключить правоприменительных требований для всех пользователей с определенным Outlook в назначенной политике почтовых ящиков, запустите следующие cmdlets.</span><span class="sxs-lookup"><span data-stu-id="eeea4-247">To turn off on-send compliance enforcement for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="eeea4-248">Windows</span><span class="sxs-lookup"><span data-stu-id="eeea4-248">Windows</span></span>](#tab/windows)

<span data-ttu-id="eeea4-249">Надстройки Outlook для Windows, использующие функцию проверки при отправке, должны запускаться для всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="eeea4-249">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="eeea4-250">Однако если пользователям требуется запустить надстройку для выполнения стандартов соответствия требованиям, групповой политике **Отключить отправку, если загрузка веб-расширений невозможна** необходимо присвоить значение **Включено** на каждом применяемом компьютере.</span><span class="sxs-lookup"><span data-stu-id="eeea4-250">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="eeea4-251">Чтобы установить политики почтовых ящиков, [](https://www.microsoft.com/download/details.aspx?id=49030) администраторы могут скачать средство административных шаблонов, а затем получить доступ к последним шаблонам администрирования, заведя редактор локальной групповой политики **gpedit.msc**.</span><span class="sxs-lookup"><span data-stu-id="eeea4-251">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy Editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="eeea4-252">Действия политики</span><span class="sxs-lookup"><span data-stu-id="eeea4-252">What the policy does</span></span>

<span data-ttu-id="eeea4-253">Для соответствия требованиям администраторам может потребоваться отключить возможность отправки сообщения или элементов собрания пользователями, пока не станет доступна для запуска последняя версия надстройки с функцией проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-253">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="eeea4-254">Администраторы должны включить групповую политику **Отключить отправку, если загрузка веб-расширений невозможна**, чтобы все надстройки обновлялись из службы Exchange и были доступны для проверки того, что каждое сообщение или элемент собрания соответствует ожидаемым правилам и нормативным требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-254">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="eeea4-255">Состояние политики</span><span class="sxs-lookup"><span data-stu-id="eeea4-255">Policy status</span></span>|<span data-ttu-id="eeea4-256">Результат</span><span class="sxs-lookup"><span data-stu-id="eeea4-256">Result</span></span>|
|---|---|
|<span data-ttu-id="eeea4-257">Отключено</span><span class="sxs-lookup"><span data-stu-id="eeea4-257">Disabled</span></span>|<span data-ttu-id="eeea4-258">Загруженные в настоящее время манифесты надстройок (не обязательно последних версий) запускаются при отправке сообщений или элементов собраний.</span><span class="sxs-lookup"><span data-stu-id="eeea4-258">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="eeea4-259">Это состояние или поведение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="eeea4-259">This is the default status/behavior.</span></span>|
|<span data-ttu-id="eeea4-260">Включен</span><span class="sxs-lookup"><span data-stu-id="eeea4-260">Enabled</span></span>|<span data-ttu-id="eeea4-261">После скачивания последних манифестов надстройки по отправке из Exchange надстройки запускаются на отправленных сообщениях или собраниях.</span><span class="sxs-lookup"><span data-stu-id="eeea4-261">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="eeea4-262">В противном случае отправка блокируется.</span><span class="sxs-lookup"><span data-stu-id="eeea4-262">Otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="eeea4-263">Управление политикой проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-263">Manage the on-send policy</span></span>

<span data-ttu-id="eeea4-264">По умолчанию политика проверки при отправке отключена.</span><span class="sxs-lookup"><span data-stu-id="eeea4-264">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="eeea4-265">Администраторы могут включить политику проверки при отправке, присвоив параметру групповой политики пользователя **Отключить отправку, если загрузка веб-расширений невозможна** значение **Включено**.</span><span class="sxs-lookup"><span data-stu-id="eeea4-265">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="eeea4-266">Чтобы отключить политику для пользователя, администратору следует присвоить ей значение **Отключено**.</span><span class="sxs-lookup"><span data-stu-id="eeea4-266">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="eeea4-267">Чтобы управлять этим параметром политики, можно выполнить указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="eeea4-267">To manage this policy setting, you can do the following.</span></span>

1. <span data-ttu-id="eeea4-268">Скачайте последнее [средство административных шаблонов](https://www.microsoft.com/download/details.aspx?id=49030).</span><span class="sxs-lookup"><span data-stu-id="eeea4-268">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="eeea4-269">Откройте редактор локальной групповой политики **(gpedit.msc).**</span><span class="sxs-lookup"><span data-stu-id="eeea4-269">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="eeea4-270">Выберите **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span><span class="sxs-lookup"><span data-stu-id="eeea4-270">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="eeea4-271">Выберите параметр **Отключить отправку, если загрузка веб-расширений невозможна**.</span><span class="sxs-lookup"><span data-stu-id="eeea4-271">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="eeea4-272">Откройте ссылку для изменения параметра политики.</span><span class="sxs-lookup"><span data-stu-id="eeea4-272">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="eeea4-273">В диалоговом окне **Отключить отправку, если загрузка веб-расширений невозможна** выберите нужный параметр (**Включено** или **Отключено**) и нажмите кнопку **OK** или **Применить**, чтобы применить обновление.</span><span class="sxs-lookup"><span data-stu-id="eeea4-273">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="eeea4-274">Mac</span><span class="sxs-lookup"><span data-stu-id="eeea4-274">Mac</span></span>](#tab/unix)

<span data-ttu-id="eeea4-275">Надстройки Outlook для Mac, использующие функцию проверки при отправке, должны запускаться у всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="eeea4-275">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="eeea4-276">Однако если пользователям требуется запустить надстройку для выполнения стандартов соответствия требованиям, необходимо применить следующий параметр почтовых ящиков на компьютере каждого пользователя.</span><span class="sxs-lookup"><span data-stu-id="eeea4-276">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="eeea4-277">Этот параметр или ключ совместим с CFPreference, то есть его можно установить, используя программное обеспечение для управления предприятием для Mac, например Jamf Pro.</span><span class="sxs-lookup"><span data-stu-id="eeea4-277">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

|||
|:---|:---|
|<span data-ttu-id="eeea4-278">**Домен**</span><span class="sxs-lookup"><span data-stu-id="eeea4-278">**Domain**</span></span>|<span data-ttu-id="eeea4-279">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-279">com.microsoft.outlook</span></span>|
|<span data-ttu-id="eeea4-280">**Ключ**</span><span class="sxs-lookup"><span data-stu-id="eeea4-280">**Key**</span></span>|<span data-ttu-id="eeea4-281">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="eeea4-281">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="eeea4-282">**Тип данных**</span><span class="sxs-lookup"><span data-stu-id="eeea4-282">**DataType**</span></span>|<span data-ttu-id="eeea4-283">Логический</span><span class="sxs-lookup"><span data-stu-id="eeea4-283">Boolean</span></span>|
|<span data-ttu-id="eeea4-284">**Возможные значения**</span><span class="sxs-lookup"><span data-stu-id="eeea4-284">**Possible values**</span></span>|<span data-ttu-id="eeea4-285">false (по умолчанию)</span><span class="sxs-lookup"><span data-stu-id="eeea4-285">false (default)</span></span><br><span data-ttu-id="eeea4-286">true</span><span class="sxs-lookup"><span data-stu-id="eeea4-286">true</span></span>|
|<span data-ttu-id="eeea4-287">**Доступность**</span><span class="sxs-lookup"><span data-stu-id="eeea4-287">**Availability**</span></span>|<span data-ttu-id="eeea4-288">16.27</span><span class="sxs-lookup"><span data-stu-id="eeea4-288">16.27</span></span>|
|<span data-ttu-id="eeea4-289">**Примечания**</span><span class="sxs-lookup"><span data-stu-id="eeea4-289">**Comments**</span></span>|<span data-ttu-id="eeea4-290">Этот ключ создает политику onSendMailbox.</span><span class="sxs-lookup"><span data-stu-id="eeea4-290">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="eeea4-291">Для чего служит параметр</span><span class="sxs-lookup"><span data-stu-id="eeea4-291">What the setting does</span></span>

<span data-ttu-id="eeea4-292">Для соответствия требованиям администраторам может потребоваться отключить возможность отправки сообщения или элементов собрания пользователями, пока не станет доступна для запуска последняя версия надстроек с функцией проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-292">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="eeea4-293">Администраторы должны включить ключ **OnSendAddinsWaitForLoad**, чтобы все надстройки обновлялись из службы Exchange и были доступны для проверки того, что каждое сообщение или элемент собрания соответствует ожидаемым правилам и нормативным требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-293">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="eeea4-294">Состояние ключа</span><span class="sxs-lookup"><span data-stu-id="eeea4-294">Key's state</span></span>|<span data-ttu-id="eeea4-295">Результат</span><span class="sxs-lookup"><span data-stu-id="eeea4-295">Result</span></span>|
|---|---|
|<span data-ttu-id="eeea4-296">false</span><span class="sxs-lookup"><span data-stu-id="eeea4-296">false</span></span>|<span data-ttu-id="eeea4-297">Загруженные в настоящее время манифесты надстройок (не обязательно последних версий) запускаются при отправке сообщений или элементов собраний.</span><span class="sxs-lookup"><span data-stu-id="eeea4-297">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="eeea4-298">Это состояние по умолчанию/поведение.</span><span class="sxs-lookup"><span data-stu-id="eeea4-298">This is the default state/behavior.</span></span>|
|<span data-ttu-id="eeea4-299">true</span><span class="sxs-lookup"><span data-stu-id="eeea4-299">true</span></span>|<span data-ttu-id="eeea4-300">После скачивания последних манифестов надстройки по отправке из Exchange надстройки запускаются на отправленных сообщениях или собраниях.</span><span class="sxs-lookup"><span data-stu-id="eeea4-300">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="eeea4-301">В противном случае отправка блокируется, а кнопка **Отправка** отключена.</span><span class="sxs-lookup"><span data-stu-id="eeea4-301">Otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="eeea4-302">Сценарии проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-302">On-send feature scenarios</span></span>

<span data-ttu-id="eeea4-303">Ниже представлены поддерживаемые и неподдерживаемые сценарии для надстроек, использующих функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-303">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="eeea4-304">В почтовом ящике пользователя включена функция проверки при отправке, но не установлено ни одной надстройки</span><span class="sxs-lookup"><span data-stu-id="eeea4-304">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="eeea4-305">В этом случае пользователь сможет отправлять сообщение или элементы собрания без запуска надстроек.</span><span class="sxs-lookup"><span data-stu-id="eeea4-305">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="eeea4-306">В почтовом ящике пользователя включена функция проверки при отправке, а также установлены и включены надстройки, поддерживающие эту функцию</span><span class="sxs-lookup"><span data-stu-id="eeea4-306">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="eeea4-307">При отправке будут запускаться надстройки, которые разрешат или заблокируют отправку.</span><span class="sxs-lookup"><span data-stu-id="eeea4-307">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="eeea4-308">Делегирование для почтовых ящиков, при котором у почтового ящика 1 есть разрешения на полный доступ к почтовому ящику 2</span><span class="sxs-lookup"><span data-stu-id="eeea4-308">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="eeea4-309">Веб-браузер (классическая версия Outlook)</span><span class="sxs-lookup"><span data-stu-id="eeea4-309">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="eeea4-310">Сценарий</span><span class="sxs-lookup"><span data-stu-id="eeea4-310">Scenario</span></span>|<span data-ttu-id="eeea4-311">Функция проверки при отправке для почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="eeea4-311">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="eeea4-312">Функция проверки при отправке для почтового ящика 2</span><span class="sxs-lookup"><span data-stu-id="eeea4-312">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="eeea4-313">Веб-сеанс Outlook (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="eeea4-313">Outlook web session (classic)</span></span>|<span data-ttu-id="eeea4-314">Результат</span><span class="sxs-lookup"><span data-stu-id="eeea4-314">Result</span></span>|<span data-ttu-id="eeea4-315">Поддержка</span><span class="sxs-lookup"><span data-stu-id="eeea4-315">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="eeea4-316">1</span><span class="sxs-lookup"><span data-stu-id="eeea4-316">1</span></span>|<span data-ttu-id="eeea4-317">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-317">Enabled</span></span>|<span data-ttu-id="eeea4-318">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-318">Enabled</span></span>|<span data-ttu-id="eeea4-319">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="eeea4-319">New session</span></span>|<span data-ttu-id="eeea4-320">Почтовый ящик 1 не может отправлять сообщение или элементы собраний из почтового ящика 2.</span><span class="sxs-lookup"><span data-stu-id="eeea4-320">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="eeea4-p135">В настоящее время не поддерживается. В качестве обходного решения используйте сценарий 3.</span><span class="sxs-lookup"><span data-stu-id="eeea4-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="eeea4-323">2</span><span class="sxs-lookup"><span data-stu-id="eeea4-323">2</span></span>|<span data-ttu-id="eeea4-324">Отключена</span><span class="sxs-lookup"><span data-stu-id="eeea4-324">Disabled</span></span>|<span data-ttu-id="eeea4-325">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-325">Enabled</span></span>|<span data-ttu-id="eeea4-326">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="eeea4-326">New session</span></span>|<span data-ttu-id="eeea4-327">Почтовый ящик 1 не может отправлять сообщение или элементы собраний из почтового ящика 2.</span><span class="sxs-lookup"><span data-stu-id="eeea4-327">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="eeea4-p136">В настоящее время не поддерживается. В качестве обходного решения используйте сценарий 3.</span><span class="sxs-lookup"><span data-stu-id="eeea4-p136">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="eeea4-330">3</span><span class="sxs-lookup"><span data-stu-id="eeea4-330">3</span></span>|<span data-ttu-id="eeea4-331">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-331">Enabled</span></span>|<span data-ttu-id="eeea4-332">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-332">Enabled</span></span>|<span data-ttu-id="eeea4-333">Тот же сеанс</span><span class="sxs-lookup"><span data-stu-id="eeea4-333">Same session</span></span>|<span data-ttu-id="eeea4-334">Проверка при отправке выполняется для почтового ящика 1, которому назначены надстройки, поддерживающие эту функцию.</span><span class="sxs-lookup"><span data-stu-id="eeea4-334">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="eeea4-335">Поддерживается.</span><span class="sxs-lookup"><span data-stu-id="eeea4-335">Supported.</span></span>|
|<span data-ttu-id="eeea4-336">4 </span><span class="sxs-lookup"><span data-stu-id="eeea4-336">4</span></span>|<span data-ttu-id="eeea4-337">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-337">Enabled</span></span>|<span data-ttu-id="eeea4-338">Отключена</span><span class="sxs-lookup"><span data-stu-id="eeea4-338">Disabled</span></span>|<span data-ttu-id="eeea4-339">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="eeea4-339">New session</span></span>|<span data-ttu-id="eeea4-340">Надстройки с функцией проверки при отправке не запускаются; отправка сообщения или элемента собрания.</span><span class="sxs-lookup"><span data-stu-id="eeea4-340">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="eeea4-341">Поддерживается.</span><span class="sxs-lookup"><span data-stu-id="eeea4-341">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="eeea4-342">Веб-браузер (современная версия Outlook), Windows, Mac</span><span class="sxs-lookup"><span data-stu-id="eeea4-342">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="eeea4-343">Чтобы внедрить функцию проверки при отправке, администраторы должны включить политику для обоих почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="eeea4-343">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="eeea4-344">Сведения о том, как поддерживать делегированный доступ в надстройке, см. в статье [Включение сценариев делегированного доступа в надстройке Outlook](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="eeea4-344">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="eeea4-345">Группа 1 — это современный почтовый ящик группы, а почтовый ящик пользователя 1 является участником группы 1</span><span class="sxs-lookup"><span data-stu-id="eeea4-345">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="eeea4-346">Сценарий</span><span class="sxs-lookup"><span data-stu-id="eeea4-346">Scenario</span></span>|<span data-ttu-id="eeea4-347">Политика проверки при отправке для почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="eeea4-347">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="eeea4-348">Включены ли надстройки, поддерживающие проверку сообщений при отправке?</span><span class="sxs-lookup"><span data-stu-id="eeea4-348">On-send add-ins enabled?</span></span>|<span data-ttu-id="eeea4-349">Действие почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="eeea4-349">Mailbox 1 action</span></span>|<span data-ttu-id="eeea4-350">Результат</span><span class="sxs-lookup"><span data-stu-id="eeea4-350">Result</span></span>|<span data-ttu-id="eeea4-351">Поддержка</span><span class="sxs-lookup"><span data-stu-id="eeea4-351">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="eeea4-352">1</span><span class="sxs-lookup"><span data-stu-id="eeea4-352">1</span></span>|<span data-ttu-id="eeea4-353">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-353">Enabled</span></span>|<span data-ttu-id="eeea4-354">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-354">Yes</span></span>|<span data-ttu-id="eeea4-355">Почтовый ящик 1 создает новое сообщение или собрание для группы 1.</span><span class="sxs-lookup"><span data-stu-id="eeea4-355">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="eeea4-356">В случае отправки запускаются надстройки, поддерживающие проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-356">On-send add-ins run during send.</span></span>|<span data-ttu-id="eeea4-357">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-357">Yes</span></span>|
|<span data-ttu-id="eeea4-358">2</span><span class="sxs-lookup"><span data-stu-id="eeea4-358">2</span></span>|<span data-ttu-id="eeea4-359">Включена</span><span class="sxs-lookup"><span data-stu-id="eeea4-359">Enabled</span></span>|<span data-ttu-id="eeea4-360">Да</span><span class="sxs-lookup"><span data-stu-id="eeea4-360">Yes</span></span>|<span data-ttu-id="eeea4-361">Почтовый ящик 1 создает новое сообщение или собрание для группы 1 в окне этой группы в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="eeea4-361">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="eeea4-362">В случае отправки не запускаются надстройки, поддерживающие проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-362">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="eeea4-363">В настоящее время не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="eeea4-363">Not currently supported.</span></span> <span data-ttu-id="eeea4-364">В качестве обходного решения используйте сценарий 1.</span><span class="sxs-lookup"><span data-stu-id="eeea4-364">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="eeea4-365">Включен почтовый ящик пользователя с функцией или политикой проверки при отправке, установлены и включены надстройки, поддерживающие эту функцию, а также включен автономный режим</span><span class="sxs-lookup"><span data-stu-id="eeea4-365">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="eeea4-366">Надстройки, поддерживающие проверку при отправке, запускаются в соответствии с сетевым состоянием пользователя, внутреннего сервера надстройки и Exchange.</span><span class="sxs-lookup"><span data-stu-id="eeea4-366">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="eeea4-367">Состояние пользователя</span><span class="sxs-lookup"><span data-stu-id="eeea4-367">User's state</span></span>

<span data-ttu-id="eeea4-368">Надстройки, поддерживающие проверку сообщений при отправке, будут запускаться при отправке, если пользователь в сети.</span><span class="sxs-lookup"><span data-stu-id="eeea4-368">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="eeea4-369">В автономном режиме такие надстройки не будут запускаться при отправке, а сообщение или элемент собрания не будет отправлен.</span><span class="sxs-lookup"><span data-stu-id="eeea4-369">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="eeea4-370">Состояние внутреннего сервера надстройки</span><span class="sxs-lookup"><span data-stu-id="eeea4-370">Add-in backend's state</span></span>

<span data-ttu-id="eeea4-371">Надстройка, поддерживающая проверку при отправке, будет запускаться, если ее внутренний сервер подключен к сети и доступен.</span><span class="sxs-lookup"><span data-stu-id="eeea4-371">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="eeea4-372">Если внутренний сервер находится в автономном режиме, отправка отключена.</span><span class="sxs-lookup"><span data-stu-id="eeea4-372">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="eeea4-373">Состояние Exchange</span><span class="sxs-lookup"><span data-stu-id="eeea4-373">Exchange's state</span></span>

<span data-ttu-id="eeea4-374">Надстройки, поддерживающие проверку сообщений при отправке, будут запускаться при отправке, если сервер Exchange подключен к сети и доступен.</span><span class="sxs-lookup"><span data-stu-id="eeea4-374">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="eeea4-375">Если надстройке с функцией проверки при отправке недоступна служба Exchange и включена соответствующая политика или командлет, отправка отключена.</span><span class="sxs-lookup"><span data-stu-id="eeea4-375">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="eeea4-376">На компьютерах Mac в любом автономном состоянии кнопка **Отправить** (или **Отправить обновление** для существующих собраний) отключена, и отображается уведомление, что в организации не разрешено отправлять сообщения, если пользователь не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="eeea4-376">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="eeea4-377">Пользователь может изменять элемент во время работы надстройки при отправке</span><span class="sxs-lookup"><span data-stu-id="eeea4-377">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="eeea4-378">В то время как надстройки при отправке обрабатывают элемент, пользователь может изменить элемент, добавив, например, неправильный текст или вложения.</span><span class="sxs-lookup"><span data-stu-id="eeea4-378">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="eeea4-379">Если вы хотите запретить пользователю изменять элемент во время обработки надстройки при отправке, вы можете реализовать обходное решение с помощью диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="eeea4-379">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="eeea4-380">Это обходное решение можно использовать в Outlook в Интернете (классический), Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="eeea4-380">This workaround can be used in Outlook on the web (classic), Windows, and Mac.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eeea4-381">Современный outlook в Интернете. Чтобы предотвратить редактирование элемента во время обработки надстройки при отправке, необходимо установить флаг *OnSendAddinsEnabled,* как описано в надстройке `true` [Install Outlook,](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) которая используется в разделе отправки в начале этой статьи.</span><span class="sxs-lookup"><span data-stu-id="eeea4-381">Modern Outlook on the web: To prevent the user from editing the item while your add-in is processing on send, you should set the *OnSendAddinsEnabled* flag to `true` as described in the [Install Outlook add-ins that use on-send](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) section earlier in this article.</span></span>

<span data-ttu-id="eeea4-382">В обработнике отправки:</span><span class="sxs-lookup"><span data-stu-id="eeea4-382">In your on-send handler:</span></span>

1. <span data-ttu-id="eeea4-383">Вызов [displayDialogAsync,](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) чтобы открыть диалоговое окно, чтобы щелчки мыши и клавиши были отключены.</span><span class="sxs-lookup"><span data-stu-id="eeea4-383">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="eeea4-384">Чтобы получить это поведение в классическом Outlook в Интернете, необходимо настроить свойство [displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) в `true` `options` параметре `displayDialogAsync` вызова.</span><span class="sxs-lookup"><span data-stu-id="eeea4-384">To get this behavior in classic Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="eeea4-385">Реализация обработки элемента.</span><span class="sxs-lookup"><span data-stu-id="eeea4-385">Implement processing of the item.</span></span>
1. <span data-ttu-id="eeea4-386">Закройте диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="eeea4-386">Close the dialog.</span></span> <span data-ttu-id="eeea4-387">Кроме того, обработать, что произойдет, если пользователь закрывает диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="eeea4-387">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="eeea4-388">Примеры кода</span><span class="sxs-lookup"><span data-stu-id="eeea4-388">Code examples</span></span>

<span data-ttu-id="eeea4-389">В приведенных ниже примерах кода показано, как создать простую надстройку, поддерживающую проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-389">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="eeea4-390">Скачать код, на котором основаны эти примеры, можно на странице [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span><span class="sxs-lookup"><span data-stu-id="eeea4-390">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="eeea4-391">Если вы используете диалоговое окно с событием отправки, не забудьте закрыть диалоговое окно перед завершением события.</span><span class="sxs-lookup"><span data-stu-id="eeea4-391">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="eeea4-392">Манифест, переопределение версии и событие</span><span class="sxs-lookup"><span data-stu-id="eeea4-392">Manifest, version override, and event</span></span>

<span data-ttu-id="eeea4-393">Пример кода [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) включает два манифеста:</span><span class="sxs-lookup"><span data-stu-id="eeea4-393">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="eeea4-394">`Contoso Message Body Checker.xml` показывает, как проверить текст сообщения на наличие запрещенных слов или конфиденциальной информации при отправке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-394">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="eeea4-395">`Contoso Subject and CC Checker.xml` показывает, как при отправке добавить получателя в строку "Копия" и проверить, включает ли сообщение строку темы.</span><span class="sxs-lookup"><span data-stu-id="eeea4-395">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="eeea4-396">В файле манифеста `Contoso Message Body Checker.xml` указываются файл и имя функции, которую следует вызывать при возникновении события `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-396">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="eeea4-397">Операция выполняется синхронно.</span><span class="sxs-lookup"><span data-stu-id="eeea4-397">The operation runs synchronously.</span></span>

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
> <span data-ttu-id="eeea4-398">Если вы используете Visual Studio 2019 для разработки надстройки при отправке, вы можете получить предупреждение о проверке, как следующее: "Это недействительный xsi:type http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events '". Чтобы обойти это, вам потребуется более новая версия MailAppVersionOverridesV1_1.xsd, которая была предоставлена в качестве gist GitHub в блоге об этом [предупреждении](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span><span class="sxs-lookup"><span data-stu-id="eeea4-398">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="eeea4-399">Для файла манифеста `Contoso Subject and CC Checker.xml` в приведенном ниже примере показаны файл и имя функции, вызываемой при возникновении события отправки.</span><span class="sxs-lookup"><span data-stu-id="eeea4-399">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

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

<span data-ttu-id="eeea4-400">Для API проверки при отправке требуется узел `VersionOverrides v1_1`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-400">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="eeea4-401">Ниже показано, как добавить узел `VersionOverrides` в манифест.</span><span class="sxs-lookup"><span data-stu-id="eeea4-401">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="eeea4-402">Дополнительные сведения см. в указанных ниже статьях.</span><span class="sxs-lookup"><span data-stu-id="eeea4-402">For more information, see the following:</span></span>
> - [<span data-ttu-id="eeea4-403">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-403">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="eeea4-404">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="eeea4-404">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="eeea4-405">Объекты `Event` и `item`, методы `body.getAsync` и `body.setAsync`</span><span class="sxs-lookup"><span data-stu-id="eeea4-405">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="eeea4-406">Чтобы получить доступ к выбранному в данный момент сообщению или элементу собрания (в этом примере — к новому сообщению), используйте пространство имен `Office.context.mailbox.item`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-406">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="eeea4-407">Функция проверки при отправке автоматически передает событие `ItemSend` функции, указанной в манифесте (в данном случае это функция `validateBody`).</span><span class="sxs-lookup"><span data-stu-id="eeea4-407">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

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

<span data-ttu-id="eeea4-408">Функция `validateBody` возвращает текущий текст в заданном формате (HTML) и передает нужный объект события `ItemSend` в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="eeea4-408">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="eeea4-409">Помимо метода `getAsync`, объект `Body` также предоставляет метод `setAsync`, с помощью которого вы можете заменить текст сообщения на указанный.</span><span class="sxs-lookup"><span data-stu-id="eeea4-409">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="eeea4-410">Дополнительные сведения см. в статьях [Объект Event](/javascript/api/office/office.addincommands.event) и [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="eeea4-410">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="eeea4-411">Объект `NotificationMessages` и метод `event.completed`</span><span class="sxs-lookup"><span data-stu-id="eeea4-411">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="eeea4-412">Функция `checkBodyOnlyOnSendCallBack` использует регулярное выражение, чтобы определить, содержит ли текст сообщения слова, подлежащие блокировке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-412">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="eeea4-413">Если она обнаруживает слово, совпадающие с каким-либо элементом из массива запрещенных слов, отправка сообщения блокируется, а отправитель получает уведомление на панели информации.</span><span class="sxs-lookup"><span data-stu-id="eeea4-413">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="eeea4-414">Для этого в ней используется свойство `notificationMessages` объекта `Item` для возврата объекта `NotificationMessages`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-414">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="eeea4-415">После этого она добавляет уведомление к элементу, вызывая метод `addAsync`, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="eeea4-415">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

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

<span data-ttu-id="eeea4-416">Ниже перечислены параметры метода `addAsync`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-416">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="eeea4-417">`NoSend`. Строка, представляющая собой заданный разработчиком ключ для ссылки на сообщение уведомления.</span><span class="sxs-lookup"><span data-stu-id="eeea4-417">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="eeea4-418">С его помощью вы сможете изменить это сообщение позже.</span><span class="sxs-lookup"><span data-stu-id="eeea4-418">You can use it to modify this message later.</span></span> <span data-ttu-id="eeea4-419">Клавиша не может быть длиннее 32 символов.</span><span class="sxs-lookup"><span data-stu-id="eeea4-419">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="eeea4-420">`type`. Одно из свойств параметра объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="eeea4-420">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="eeea4-421">Представляет тип сообщения. Типы соответствуют значениям перечисления [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype).</span><span class="sxs-lookup"><span data-stu-id="eeea4-421">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="eeea4-422">Допустимые значения: индикатор хода выполнения, информационное сообщение и сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-422">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="eeea4-423">В этом примере в свойстве `type` указано сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="eeea4-423">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="eeea4-424">`message`. Одно из свойств параметра объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="eeea4-424">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="eeea4-425">В этом примере `message` — это текст сообщения уведомления.</span><span class="sxs-lookup"><span data-stu-id="eeea4-425">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="eeea4-426">Чтобы сообщить о завершении надстройкой обработки события `ItemSend`, активированного операцией отправки, вызовите метод `event.completed({allowEvent:Boolean})`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-426">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="eeea4-427">Свойство `allowEvent` является логическим.</span><span class="sxs-lookup"><span data-stu-id="eeea4-427">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="eeea4-428">Если задано значение `true`, отправка разрешается.</span><span class="sxs-lookup"><span data-stu-id="eeea4-428">If set to `true`, send is allowed.</span></span> <span data-ttu-id="eeea4-429">Если задано значение `false`, отправка письма блокируется.</span><span class="sxs-lookup"><span data-stu-id="eeea4-429">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="eeea4-430">Дополнительные сведения см. в статьях [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [completed](/javascript/api/office/office.addincommands.event).</span><span class="sxs-lookup"><span data-stu-id="eeea4-430">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="eeea4-431">Методы `replaceAsync`, `removeAsync` и `getAllAsync`</span><span class="sxs-lookup"><span data-stu-id="eeea4-431">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="eeea4-432">Помимо метода `addAsync`, объект `NotificationMessages` также включает методы `replaceAsync`, `removeAsync` и `getAllAsync`.</span><span class="sxs-lookup"><span data-stu-id="eeea4-432">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="eeea4-433">Эти методы не используются в данном примере кода.</span><span class="sxs-lookup"><span data-stu-id="eeea4-433">These methods are not used in this code sample.</span></span>  <span data-ttu-id="eeea4-434">Дополнительные сведения см. в статье [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span><span class="sxs-lookup"><span data-stu-id="eeea4-434">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="eeea4-435">Код проверки строк "Тема" и "Копия"</span><span class="sxs-lookup"><span data-stu-id="eeea4-435">Subject and CC checker code</span></span>

<span data-ttu-id="eeea4-436">В приведенном ниже примере кода показано, как при отправке сообщения добавить получателя в строку "Копия" и проверить, включает ли сообщение тему.</span><span class="sxs-lookup"><span data-stu-id="eeea4-436">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="eeea4-437">В этом примере функция проверки при отправке используется, чтобы разрешить или запретить отправку сообщения.</span><span class="sxs-lookup"><span data-stu-id="eeea4-437">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

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

<span data-ttu-id="eeea4-p156">Дополнительные сведения о том, как при отправке сообщения добавить получателя в строку "Копия" и проверить, указана ли тема сообщения, а также просмотреть доступные API, см. в [примере Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). Код сопровождается подробными комментариями.</span><span class="sxs-lookup"><span data-stu-id="eeea4-p156">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="eeea4-440">См. также</span><span class="sxs-lookup"><span data-stu-id="eeea4-440">See also</span></span>

- [<span data-ttu-id="eeea4-441">Обзор архитектуры и функций надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="eeea4-441">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="eeea4-442">Надстройка Outlook "Демонстрация команд надстройки"</span><span class="sxs-lookup"><span data-stu-id="eeea4-442">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)