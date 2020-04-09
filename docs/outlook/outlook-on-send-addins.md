---
title: Функция проверки при отправке для надстроек Outlook
description: Позволяет надстройке настраивать те или иные параметры при отправке, а также обрабатывать элемент и запрещать пользователям выполнять определенные действия.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: 017759628cd9b3716c3992f7c6631911491ca246
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185640"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="643d5-103">Функция проверки при отправке для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="643d5-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="643d5-p101">Функция проверки при отправке позволяет надстройкам Outlook настраивать те или иные параметры при отправке, а также обрабатывать сообщение или элемент собрания и запрещать пользователям выполнять определенные действия. Например, с помощью функции проверки при отправке можно сделать следующее:</span><span class="sxs-lookup"><span data-stu-id="643d5-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="643d5-106">запретить пользователю отправлять конфиденциальную информацию или оставлять строку темы пустой;</span><span class="sxs-lookup"><span data-stu-id="643d5-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="643d5-107">добавить определенного получателя в строку "Копия" в сообщениях или в строку "Необязательные получатели" в собраниях.</span><span class="sxs-lookup"><span data-stu-id="643d5-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

> [!NOTE]
> <span data-ttu-id="643d5-108">Функция проверки при отправке в настоящее время поддерживается для Outlook в Интернете в Exchange Online (Office 365), локальной среде Exchange 2016 (накопительный пакет обновления 6 или более поздняя версия) и локальной среде Exchange 2019 (накопительный пакет обновления 1 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="643d5-108">The on-send feature is currently supported for Outlook on the web in Exchange Online (Office 365), Exchange 2016 on-premises (Cumulative Update 6 or later), and Exchange 2019 on-premises (Cumulative Update 1 or later).</span></span> <span data-ttu-id="643d5-109">Эта функция также доступна в последних сборках Outlook для Windows и Mac, подключенных к Exchange Online (Office 365).</span><span class="sxs-lookup"><span data-stu-id="643d5-109">This feature is also available in the latest Outlook builds on Windows and Mac, connected to Exchange Online (Office 365).</span></span> <span data-ttu-id="643d5-110">Эта функция появилась в наборе требований 1,8 (Дополнительные сведения см. в статье [Current Server and Client Support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) ).</span><span class="sxs-lookup"><span data-stu-id="643d5-110">The feature was introduced in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="643d5-111">Надстройки, использующие функцию отправки при отправке, не разрешены в [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="643d5-111">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="643d5-112">Функция проверки при отправке не использует пользовательский интерфейс и активируется событием типа `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="643d5-112">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="643d5-113">Сведения об ограничениях, связанных с функцией проверки при отправке, см. в разделе [Ограничения](#limitations) далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="643d5-113">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="643d5-114">Как работает функция проверки при отправке?</span><span class="sxs-lookup"><span data-stu-id="643d5-114">How does the on-send feature work?</span></span>

<span data-ttu-id="643d5-115">С помощью функции проверки при отправке вы можете создать надстройку Outlook, задействующую синхронное событие `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="643d5-115">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="643d5-116">Это событие возникает, когда пользователь нажимает кнопку **Отправить** (или **Отправить обновление** для существующих собраний). С его помощью можно блокировать отправку элемента, не прошедшего проверку.</span><span class="sxs-lookup"><span data-stu-id="643d5-116">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="643d5-117">Например, когда пользователь вызывает событие отправки сообщения, надстройка Outlook благодаря функции проверки при отправке может следующее:</span><span class="sxs-lookup"><span data-stu-id="643d5-117">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="643d5-118">прочитать и проверить содержимое письма;</span><span class="sxs-lookup"><span data-stu-id="643d5-118">Read and validate the email message contents</span></span>
- <span data-ttu-id="643d5-119">проверить, содержит ли сообщение строку темы;</span><span class="sxs-lookup"><span data-stu-id="643d5-119">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="643d5-120">задать заранее определенного получателя.</span><span class="sxs-lookup"><span data-stu-id="643d5-120">Set a predetermined recipient</span></span>

<span data-ttu-id="643d5-121">Проверка выполняется на клиентской стороне Outlook, когда инициируется событие Send, и надстройка занимает до истечения времени до 5 минут. Если проверка завершается неудачно, отправка элемента блокируется, и на панели информации отображается сообщение об ошибке, предлагающее пользователю выполнить действие.</span><span class="sxs-lookup"><span data-stu-id="643d5-121">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

<span data-ttu-id="643d5-122">На приведенном ниже снимке экрана показана панель информации, в которой пользователю предлагается добавить тему.</span><span class="sxs-lookup"><span data-stu-id="643d5-122">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![Снимок экрана с сообщением об ошибке, в котором пользователю предлагается ввести отсутствующую строку темы](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="643d5-124">На приведенном ниже снимке экрана показана панель информации, уведомляющая отправителя о том, что обнаружены слова, подлежащие блокировке.</span><span class="sxs-lookup"><span data-stu-id="643d5-124">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![Снимок экрана с сообщением об ошибке, которое сообщает пользователю о том, что обнаружены слова, подлежащие блокировке](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="643d5-126">Ограничения</span><span class="sxs-lookup"><span data-stu-id="643d5-126">Limitations</span></span>

<span data-ttu-id="643d5-127">В настоящее время на функцию проверки при отправке действуют перечисленные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="643d5-127">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="643d5-128">**AppSource**. Надстройки Outlook, в которых используется функция проверки при отправке, невозможно публиковать в [AppSource](https://appsource.microsoft.com), так как они не проходят проверку AppSource.</span><span class="sxs-lookup"><span data-stu-id="643d5-128">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="643d5-129">Надстройки, использующие функцию проверки при отправке, должны разворачиваться администраторами.</span><span class="sxs-lookup"><span data-stu-id="643d5-129">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="643d5-130">**Манифест**. Для каждой надстройки поддерживается только одно событие `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="643d5-130">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="643d5-131">Если манифест содержит несколько событий `ItemSend`, он не пройдет проверку.</span><span class="sxs-lookup"><span data-stu-id="643d5-131">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="643d5-p106">**Производительность**. &ndash;Многочисленные случаи приема-передачи пакетов на веб-сервере, где размещается надстройка, могут повлиять на ее производительность. Учитывайте влияние на производительность при создании надстроек, требующих выполнения нескольких операций с сообщениями или собраниями.</span><span class="sxs-lookup"><span data-stu-id="643d5-p106">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="643d5-134">**Отправить позже** (только для Mac). Если имеются надстройки, использующие проверку при отправке, функция **Отправить позже** будет недоступна.</span><span class="sxs-lookup"><span data-stu-id="643d5-134">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="643d5-135">Ограничения на типы и режимы почтовых ящиков</span><span class="sxs-lookup"><span data-stu-id="643d5-135">Mailbox type/mode limitations</span></span>

<span data-ttu-id="643d5-136">Функция проверки при отправке поддерживается только для почтовых ящиков пользователей в Outlook в Интернете, Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="643d5-136">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="643d5-137">В настоящее время она не поддерживается для указанных ниже типов и режимов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-137">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="643d5-138">Общие почтовые ящики\*</span><span class="sxs-lookup"><span data-stu-id="643d5-138">Shared mailboxes\*</span></span>
- <span data-ttu-id="643d5-139">почтовые ящики групп;</span><span class="sxs-lookup"><span data-stu-id="643d5-139">Group mailboxes</span></span>
- <span data-ttu-id="643d5-140">почтовые ящики в автономном режиме.</span><span class="sxs-lookup"><span data-stu-id="643d5-140">Offline mode</span></span>

<span data-ttu-id="643d5-141">Outlook не разрешает отправку, если для этих сценариев почтовых ящиков включена функция проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-141">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="643d5-142">Однако если пользователь отвечает на сообщение в почтовом ящике группы, надстройка, поддерживающая проверку сообщений при отправке, не запускается, а сообщение отправляется.</span><span class="sxs-lookup"><span data-stu-id="643d5-142">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="643d5-143">\*Функция On – Send должна работать с общими почтовыми ящиками или папками, если надстройка также [реализует поддержку сценариев делегированного доступа](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="643d5-143">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="643d5-144">Несколько надстроек, поддерживающих проверку сообщений при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-144">Multiple on-send add-ins</span></span>

<span data-ttu-id="643d5-145">Если установлено несколько надстроек, поддерживающих проверку сообщений при отправке, они будут запускаться в том порядке, в котором были получены из API `getAppManifestCall` или `getExtensibilityContext`.</span><span class="sxs-lookup"><span data-stu-id="643d5-145">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="643d5-146">Если первая надстройка разрешает отправку сообщения, то вторая может внести изменения, на которые первая отреагировала бы запретом отправки.</span><span class="sxs-lookup"><span data-stu-id="643d5-146">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="643d5-147">Однако первая надстройка не будет запускаться снова, если все установленные надстройки разрешат отправку.</span><span class="sxs-lookup"><span data-stu-id="643d5-147">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="643d5-148">Допустим, надстройка 1 и надстройка 2 используют функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-148">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="643d5-149">Сначала устанавливается надстройка 1, а затем — надстройка 2.</span><span class="sxs-lookup"><span data-stu-id="643d5-149">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="643d5-150">Надстройка 1 находит в сообщении слово Fabrikam, что является условием для разрешения отправки.</span><span class="sxs-lookup"><span data-stu-id="643d5-150">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="643d5-151">Однако надстройка 2 удаляет все вхождения слова Fabrikam.</span><span class="sxs-lookup"><span data-stu-id="643d5-151">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="643d5-152">Сообщение будет отправлено без единого слова Fabrikam (в связи с порядком установки надстроек 1 и 2).</span><span class="sxs-lookup"><span data-stu-id="643d5-152">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="643d5-153">Развертывание надстроек Outlook, использующих функцию проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-153">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="643d5-154">Рекомендуем, чтобы развертывание надстроек Outlook, использующих функцию проверки при отправке, выполняли администраторы.</span><span class="sxs-lookup"><span data-stu-id="643d5-154">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="643d5-155">Администратор должен убедиться, что такая надстройка:</span><span class="sxs-lookup"><span data-stu-id="643d5-155">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="643d5-156">всегда присутствует при открытии создаваемого элемента (для электронной почты: создание сообщений, ответ и пересылка);</span><span class="sxs-lookup"><span data-stu-id="643d5-156">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="643d5-157">не может быть закрыта или отключена пользователем.</span><span class="sxs-lookup"><span data-stu-id="643d5-157">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="643d5-158">Установка надстроек Outlook, использующих функцию проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-158">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="643d5-159">Чтобы использовать функцию проверки при отправке в Outlook, надстройки должны быть настроены для типов событий отправки.</span><span class="sxs-lookup"><span data-stu-id="643d5-159">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="643d5-160">Выберите платформу, которую нужно настроить.</span><span class="sxs-lookup"><span data-stu-id="643d5-160">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="643d5-161">Веб-браузер — классическая версия Outlook</span><span class="sxs-lookup"><span data-stu-id="643d5-161">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="643d5-162">Надстройки для Outlook в Интернете (классическая версия), использующие функцию проверки при отправке, будут запускаться у пользователей, которым назначена политика почтовых ящиков Outlook в Интернете, для флага *OnSendAddinsEnabled* которой задано значение **true**.</span><span class="sxs-lookup"><span data-stu-id="643d5-162">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="643d5-163">Чтобы установить новую надстройку, выполните приведенные ниже командлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="643d5-163">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="643d5-164">Сведения о том, как подключиться к Exchange Online с помощью удаленного сеанса PowerShell, см. в статье [Подключение к Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="643d5-164">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="643d5-165">Включение функции проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-165">Enable the on-send feature</span></span>

<span data-ttu-id="643d5-166">По умолчанию функция проверки при отправке отключена.</span><span class="sxs-lookup"><span data-stu-id="643d5-166">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="643d5-167">Администраторы могут включать эту функцию с помощью командлетов Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="643d5-167">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="643d5-168">Чтобы для всех пользователей включить надстройки, поддерживающие проверку сообщений при отправке, сделайте следующее:</span><span class="sxs-lookup"><span data-stu-id="643d5-168">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="643d5-169">Создайте политику почтовых ящиков Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="643d5-169">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="643d5-170">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-170">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="643d5-171">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-171">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="643d5-172">Включите функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-172">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="643d5-173">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="643d5-173">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="643d5-174">Включение функции проверки при отправке для группы пользователей</span><span class="sxs-lookup"><span data-stu-id="643d5-174">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="643d5-175">Чтобы включить функцию проверки при отправке для определенной группы пользователей, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="643d5-175">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="643d5-176">В этом примере администратор включает функцию проверки при отправке для надстроек Outlook в Интернете только в среде финансового отдела (Finance).</span><span class="sxs-lookup"><span data-stu-id="643d5-176">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="643d5-177">Создайте политику почтовых ящиков Outlook в Интернете для группы.</span><span class="sxs-lookup"><span data-stu-id="643d5-177">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="643d5-178">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков (дополнительные сведения см. в разделе [Ограничения на типы почтовых ящиков](#multiple-on-send-add-ins) выше в этой статье).</span><span class="sxs-lookup"><span data-stu-id="643d5-178">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="643d5-179">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-179">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="643d5-180">Включите функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-180">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="643d5-181">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="643d5-181">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="643d5-182">Дождитесь вступления политики в силу (это может занять до 60 минут) или перезапустите службы IIS.</span><span class="sxs-lookup"><span data-stu-id="643d5-182">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="643d5-183">Когда политика вступит в силу, для группы будет включена функция проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-183">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="643d5-184">Отключение функции проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-184">Disable the on-send feature</span></span>

<span data-ttu-id="643d5-185">Чтобы отключить функцию проверки при отправке для пользователя или назначить политику почтовых ящиков Outlook в Интернете, в которой не включен соответствующий флаг, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="643d5-185">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="643d5-186">В этом примере используется политика почтовых ящиков *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="643d5-186">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="643d5-187">Дополнительные сведения о том, как настроить существующие политики почтовых ящиков Outlook в Интернете с помощью командлета **Set-OwaMailboxPolicy**, см. в статье [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="643d5-187">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="643d5-188">Чтобы отключить функцию проверки при отправке для всех пользователей, которым назначена определенная политика почтовых ящиков Outlook в Интернете, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="643d5-188">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="643d5-189">Веб-браузер — современная версия Outlook</span><span class="sxs-lookup"><span data-stu-id="643d5-189">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="643d5-190">Надстройки для Outlook в Интернете (современная версия), использующие функцию проверки при отправке, должны запускаться для всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="643d5-190">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="643d5-191">Однако если пользователям требуется запустить надстройку для выполнения стандартов соответствия требованиям, для флага *OnSendAddinsEnabled* политики почтовых ящиков необходимо установить значение **true**.</span><span class="sxs-lookup"><span data-stu-id="643d5-191">However, if users are required to run the add-in to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="643d5-192">Чтобы установить новую надстройку, выполните приведенные ниже командлеты Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="643d5-192">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="643d5-193">Сведения о том, как подключиться к Exchange Online с помощью удаленного сеанса PowerShell, см. в статье [Подключение к Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span><span class="sxs-lookup"><span data-stu-id="643d5-193">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-policy"></a><span data-ttu-id="643d5-194">Включение политики проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-194">Enable the on-send policy</span></span>

<span data-ttu-id="643d5-195">По умолчанию политика проверки при отправке отключена.</span><span class="sxs-lookup"><span data-stu-id="643d5-195">By default, on-send policy is disabled.</span></span> <span data-ttu-id="643d5-196">Администраторы могут включать эту функцию с помощью командлетов Exchange Online PowerShell.</span><span class="sxs-lookup"><span data-stu-id="643d5-196">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="643d5-197">Чтобы для всех пользователей включить надстройки, поддерживающие проверку сообщений при отправке, сделайте следующее:</span><span class="sxs-lookup"><span data-stu-id="643d5-197">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="643d5-198">Создайте политику почтовых ящиков Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="643d5-198">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="643d5-199">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-199">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="643d5-200">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-200">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="643d5-201">Включите функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-201">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="643d5-202">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="643d5-202">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-policy-for-a-group-of-users"></a><span data-ttu-id="643d5-203">Включение политики проверки при отправке для группы пользователей</span><span class="sxs-lookup"><span data-stu-id="643d5-203">Enable the on-send policy for a group of users</span></span>

<span data-ttu-id="643d5-204">Чтобы включить политику проверки при отправке для определенной группы пользователей, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="643d5-204">To enable the on-send policy for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="643d5-205">В этом примере администратор включает политику проверки при отправке для надстроек Outlook в Интернете только в среде финансового отдела (Finance).</span><span class="sxs-lookup"><span data-stu-id="643d5-205">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="643d5-206">Создайте политику почтовых ящиков Outlook в Интернете для группы.</span><span class="sxs-lookup"><span data-stu-id="643d5-206">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="643d5-207">Администраторы могут использовать существующую политику, но функция проверки при отправке поддерживается только для определенных типов почтовых ящиков (дополнительные сведения см. в разделе [Ограничения на типы почтовых ящиков](#multiple-on-send-add-ins) выше в этой статье).</span><span class="sxs-lookup"><span data-stu-id="643d5-207">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="643d5-208">По умолчанию в Outlook в Интернете блокируется отправка сообщений из неподдерживаемых почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-208">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="643d5-209">Включите политику проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-209">Enable the on-send policy.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="643d5-210">Назначьте политику пользователям.</span><span class="sxs-lookup"><span data-stu-id="643d5-210">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="643d5-211">Дождитесь вступления политики в силу (это может занять до 60 минут) или перезапустите службы IIS.</span><span class="sxs-lookup"><span data-stu-id="643d5-211">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="643d5-212">Когда политика вступит в силу, для группы будет внедрена функция проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-212">When the policy takes effect, the on-send feature will be enforced for the group.</span></span>

#### <a name="disable-the-on-send-policy"></a><span data-ttu-id="643d5-213">Отключение политики проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-213">Disable the on-send policy</span></span>

<span data-ttu-id="643d5-214">Чтобы отключить политику проверки при отправке для пользователя или назначить политику почтовых ящиков Outlook в Интернете, в которой не включен соответствующий флаг, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="643d5-214">To disable the on-send policy for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="643d5-215">В этом примере используется политика почтовых ящиков *ContosoCorpOWAPolicy*.</span><span class="sxs-lookup"><span data-stu-id="643d5-215">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="643d5-216">Дополнительные сведения о том, как настроить существующие политики почтовых ящиков Outlook в Интернете с помощью командлета **Set-OwaMailboxPolicy**, см. в статье [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span><span class="sxs-lookup"><span data-stu-id="643d5-216">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="643d5-217">Чтобы отключить политику проверки при отправке для всех пользователей, которым назначена определенная политика почтовых ящиков Outlook в Интернете, выполните приведенные ниже командлеты.</span><span class="sxs-lookup"><span data-stu-id="643d5-217">To disable the on-send policy for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="643d5-218">Windows</span><span class="sxs-lookup"><span data-stu-id="643d5-218">Windows</span></span>](#tab/windows)

<span data-ttu-id="643d5-219">Надстройки Outlook для Windows, использующие функцию проверки при отправке, должны запускаться для всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="643d5-219">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="643d5-220">Однако если пользователям требуется запустить надстройку для выполнения стандартов соответствия требованиям, групповой политике **Отключить отправку, если загрузка веб-расширений невозможна** необходимо присвоить значение **Включено** на каждом применяемом компьютере.</span><span class="sxs-lookup"><span data-stu-id="643d5-220">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="643d5-221">Чтобы настроить политики почтовых ящиков, администраторы могут скачать [средство административных шаблонов](https://www.microsoft.com/download/details.aspx?id=49030) и открыть последние административные шаблоны, запустив редактор локальных групповых политик **gpedit.msc**.</span><span class="sxs-lookup"><span data-stu-id="643d5-221">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="643d5-222">Действия политики</span><span class="sxs-lookup"><span data-stu-id="643d5-222">What the policy does</span></span>

<span data-ttu-id="643d5-223">Для соответствия требованиям администраторам может потребоваться отключить возможность отправки сообщения или элементов собрания пользователями, пока не станет доступна для запуска последняя версия надстройки с функцией проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-223">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="643d5-224">Администраторы должны включить групповую политику **Отключить отправку, если загрузка веб-расширений невозможна**, чтобы все надстройки обновлялись из службы Exchange и были доступны для проверки того, что каждое сообщение или элемент собрания соответствует ожидаемым правилам и нормативным требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-224">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="643d5-225">Состояние политики</span><span class="sxs-lookup"><span data-stu-id="643d5-225">Policy status</span></span>|<span data-ttu-id="643d5-226">Результат</span><span class="sxs-lookup"><span data-stu-id="643d5-226">Result</span></span>|
|---|---|
|<span data-ttu-id="643d5-227">Отключено</span><span class="sxs-lookup"><span data-stu-id="643d5-227">Disabled</span></span>|<span data-ttu-id="643d5-228">Отправка разрешена.</span><span class="sxs-lookup"><span data-stu-id="643d5-228">Send allowed.</span></span> <span data-ttu-id="643d5-229">Сообщение или элемент собрания можно отправлять без запуска надстройки, использующей проверку при отправке, даже если она не была обновлена с помощью Exchange.</span><span class="sxs-lookup"><span data-stu-id="643d5-229">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="643d5-230">Включено</span><span class="sxs-lookup"><span data-stu-id="643d5-230">Enabled</span></span>|<span data-ttu-id="643d5-231">Отправка разрешена только в том случае, если надстройка была обновлена с помощью Exchange; в противном случае отправка блокируется.</span><span class="sxs-lookup"><span data-stu-id="643d5-231">Send allowed only when the add-in has been updated from Exchange; otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="643d5-232">Управление политикой проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-232">Manage the on-send policy</span></span>

<span data-ttu-id="643d5-233">По умолчанию политика проверки при отправке отключена.</span><span class="sxs-lookup"><span data-stu-id="643d5-233">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="643d5-234">Администраторы могут включить политику проверки при отправке, присвоив параметру групповой политики пользователя **Отключить отправку, если загрузка веб-расширений невозможна** значение **Включено**.</span><span class="sxs-lookup"><span data-stu-id="643d5-234">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="643d5-235">Чтобы отключить политику для пользователя, администратору следует присвоить ей значение **Отключено**.</span><span class="sxs-lookup"><span data-stu-id="643d5-235">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="643d5-236">Чтобы управлять этим параметром политики, можно выполнить указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="643d5-236">To manage this policy setting, you can do the following.</span></span>

1. <span data-ttu-id="643d5-237">Скачайте последнее [средство административных шаблонов](https://www.microsoft.com/download/details.aspx?id=49030).</span><span class="sxs-lookup"><span data-stu-id="643d5-237">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="643d5-238">Откройте редактор локальных групповых политик (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="643d5-238">Open the Local Group Policy editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="643d5-239">Выберите **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span><span class="sxs-lookup"><span data-stu-id="643d5-239">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="643d5-240">Выберите параметр **Отключить отправку, если загрузка веб-расширений невозможна**.</span><span class="sxs-lookup"><span data-stu-id="643d5-240">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="643d5-241">Откройте ссылку для изменения параметра политики.</span><span class="sxs-lookup"><span data-stu-id="643d5-241">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="643d5-242">В диалоговом окне **Отключить отправку, если загрузка веб-расширений невозможна** выберите нужный параметр (**Включено** или **Отключено**) и нажмите кнопку **OK** или **Применить**, чтобы применить обновление.</span><span class="sxs-lookup"><span data-stu-id="643d5-242">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="643d5-243">Mac</span><span class="sxs-lookup"><span data-stu-id="643d5-243">Mac</span></span>](#tab/unix)

<span data-ttu-id="643d5-244">Надстройки Outlook для Mac, использующие функцию проверки при отправке, должны запускаться у всех пользователей, установивших их.</span><span class="sxs-lookup"><span data-stu-id="643d5-244">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="643d5-245">Однако если пользователям требуется запустить надстройку для выполнения стандартов соответствия требованиям, необходимо применить следующий параметр почтовых ящиков на компьютере каждого пользователя.</span><span class="sxs-lookup"><span data-stu-id="643d5-245">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="643d5-246">Этот параметр или ключ совместим с CFPreference, то есть его можно установить, используя программное обеспечение для управления предприятием для Mac, например Jamf Pro.</span><span class="sxs-lookup"><span data-stu-id="643d5-246">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

|||
|:---|:---|
|<span data-ttu-id="643d5-247">**Домен**</span><span class="sxs-lookup"><span data-stu-id="643d5-247">**Domain**</span></span>|<span data-ttu-id="643d5-248">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="643d5-248">com.microsoft.outlook</span></span>|
|<span data-ttu-id="643d5-249">**Ключ**</span><span class="sxs-lookup"><span data-stu-id="643d5-249">**Key**</span></span>|<span data-ttu-id="643d5-250">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="643d5-250">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="643d5-251">**Тип данных**</span><span class="sxs-lookup"><span data-stu-id="643d5-251">**DataType**</span></span>|<span data-ttu-id="643d5-252">Логический</span><span class="sxs-lookup"><span data-stu-id="643d5-252">Boolean</span></span>|
|<span data-ttu-id="643d5-253">**Возможные значения**</span><span class="sxs-lookup"><span data-stu-id="643d5-253">**Possible values**</span></span>|<span data-ttu-id="643d5-254">false (по умолчанию)</span><span class="sxs-lookup"><span data-stu-id="643d5-254">false (default)</span></span><br><span data-ttu-id="643d5-255">true</span><span class="sxs-lookup"><span data-stu-id="643d5-255">true</span></span>|
|<span data-ttu-id="643d5-256">**Доступность**</span><span class="sxs-lookup"><span data-stu-id="643d5-256">**Availability**</span></span>|<span data-ttu-id="643d5-257">16.27</span><span class="sxs-lookup"><span data-stu-id="643d5-257">16.27</span></span>|
|<span data-ttu-id="643d5-258">**Примечания**</span><span class="sxs-lookup"><span data-stu-id="643d5-258">**Comments**</span></span>|<span data-ttu-id="643d5-259">Этот ключ создает политику onSendMailbox.</span><span class="sxs-lookup"><span data-stu-id="643d5-259">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="643d5-260">Для чего служит параметр</span><span class="sxs-lookup"><span data-stu-id="643d5-260">What the setting does</span></span>

<span data-ttu-id="643d5-261">Для соответствия требованиям администраторам может потребоваться отключить возможность отправки сообщения или элементов собрания пользователями, пока не станет доступна для запуска последняя версия надстроек с функцией проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-261">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="643d5-262">Администраторы должны включить ключ **OnSendAddinsWaitForLoad**, чтобы все надстройки обновлялись из службы Exchange и были доступны для проверки того, что каждое сообщение или элемент собрания соответствует ожидаемым правилам и нормативным требованиям при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-262">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="643d5-263">Состояние ключа</span><span class="sxs-lookup"><span data-stu-id="643d5-263">Key's state</span></span>|<span data-ttu-id="643d5-264">Результат</span><span class="sxs-lookup"><span data-stu-id="643d5-264">Result</span></span>|
|---|---|
|<span data-ttu-id="643d5-265">false</span><span class="sxs-lookup"><span data-stu-id="643d5-265">false</span></span>|<span data-ttu-id="643d5-266">Отправка разрешена.</span><span class="sxs-lookup"><span data-stu-id="643d5-266">Send allowed.</span></span> <span data-ttu-id="643d5-267">Сообщение или элемент собрания можно отправлять без запуска надстройки, использующей проверку при отправке, даже если она не была обновлена с помощью Exchange.</span><span class="sxs-lookup"><span data-stu-id="643d5-267">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="643d5-268">true</span><span class="sxs-lookup"><span data-stu-id="643d5-268">true</span></span>|<span data-ttu-id="643d5-269">Отправка разрешена только в том случае, если надстройки были обновлены с помощью Exchange; в противном случае отправка блокируется и кнопка **Отправить** отключена.</span><span class="sxs-lookup"><span data-stu-id="643d5-269">Send allowed only when add-ins have been updated from Exchange; otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="643d5-270">Сценарии проверки при отправке</span><span class="sxs-lookup"><span data-stu-id="643d5-270">On-send feature scenarios</span></span>

<span data-ttu-id="643d5-271">Ниже представлены поддерживаемые и неподдерживаемые сценарии для надстроек, использующих функцию проверки при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-271">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="643d5-272">В почтовом ящике пользователя включена функция проверки при отправке, но не установлено ни одной надстройки</span><span class="sxs-lookup"><span data-stu-id="643d5-272">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="643d5-273">В этом случае пользователь сможет отправлять сообщение или элементы собрания без запуска надстроек.</span><span class="sxs-lookup"><span data-stu-id="643d5-273">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="643d5-274">В почтовом ящике пользователя включена функция проверки при отправке, а также установлены и включены надстройки, поддерживающие эту функцию</span><span class="sxs-lookup"><span data-stu-id="643d5-274">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="643d5-275">При отправке будут запускаться надстройки, которые разрешат или заблокируют отправку.</span><span class="sxs-lookup"><span data-stu-id="643d5-275">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="643d5-276">Делегирование для почтовых ящиков, при котором у почтового ящика 1 есть разрешения на полный доступ к почтовому ящику 2</span><span class="sxs-lookup"><span data-stu-id="643d5-276">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="643d5-277">Веб-браузер (классическая версия Outlook)</span><span class="sxs-lookup"><span data-stu-id="643d5-277">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="643d5-278">Сценарий</span><span class="sxs-lookup"><span data-stu-id="643d5-278">Scenario</span></span>|<span data-ttu-id="643d5-279">Функция проверки при отправке для почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="643d5-279">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="643d5-280">Функция проверки при отправке для почтового ящика 2</span><span class="sxs-lookup"><span data-stu-id="643d5-280">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="643d5-281">Веб-сеанс Outlook (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="643d5-281">Outlook web session (classic)</span></span>|<span data-ttu-id="643d5-282">Результат</span><span class="sxs-lookup"><span data-stu-id="643d5-282">Result</span></span>|<span data-ttu-id="643d5-283">Поддержка</span><span class="sxs-lookup"><span data-stu-id="643d5-283">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="643d5-284">1,1</span><span class="sxs-lookup"><span data-stu-id="643d5-284">1</span></span>|<span data-ttu-id="643d5-285">Включено</span><span class="sxs-lookup"><span data-stu-id="643d5-285">Enabled</span></span>|<span data-ttu-id="643d5-286">Включено</span><span class="sxs-lookup"><span data-stu-id="643d5-286">Enabled</span></span>|<span data-ttu-id="643d5-287">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="643d5-287">New session</span></span>|<span data-ttu-id="643d5-288">Почтовый ящик 1 не может отправлять сообщение или элементы собраний из почтового ящика 2.</span><span class="sxs-lookup"><span data-stu-id="643d5-288">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="643d5-p133">В настоящее время не поддерживается. В качестве обходного решения используйте сценарий 3.</span><span class="sxs-lookup"><span data-stu-id="643d5-p133">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="643d5-291">2</span><span class="sxs-lookup"><span data-stu-id="643d5-291">2</span></span>|<span data-ttu-id="643d5-292">Отключена</span><span class="sxs-lookup"><span data-stu-id="643d5-292">Disabled</span></span>|<span data-ttu-id="643d5-293">Включено</span><span class="sxs-lookup"><span data-stu-id="643d5-293">Enabled</span></span>|<span data-ttu-id="643d5-294">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="643d5-294">New session</span></span>|<span data-ttu-id="643d5-295">Почтовый ящик 1 не может отправлять сообщение или элементы собраний из почтового ящика 2.</span><span class="sxs-lookup"><span data-stu-id="643d5-295">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="643d5-p134">В настоящее время не поддерживается. В качестве обходного решения используйте сценарий 3.</span><span class="sxs-lookup"><span data-stu-id="643d5-p134">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="643d5-298">4</span><span class="sxs-lookup"><span data-stu-id="643d5-298">3</span></span>|<span data-ttu-id="643d5-299">Включена</span><span class="sxs-lookup"><span data-stu-id="643d5-299">Enabled</span></span>|<span data-ttu-id="643d5-300">Включена</span><span class="sxs-lookup"><span data-stu-id="643d5-300">Enabled</span></span>|<span data-ttu-id="643d5-301">Тот же сеанс</span><span class="sxs-lookup"><span data-stu-id="643d5-301">Same session</span></span>|<span data-ttu-id="643d5-302">Проверка при отправке выполняется для почтового ящика 1, которому назначены надстройки, поддерживающие эту функцию.</span><span class="sxs-lookup"><span data-stu-id="643d5-302">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="643d5-303">Поддерживается.</span><span class="sxs-lookup"><span data-stu-id="643d5-303">Supported.</span></span>|
|<span data-ttu-id="643d5-304">4 </span><span class="sxs-lookup"><span data-stu-id="643d5-304">4</span></span>|<span data-ttu-id="643d5-305">Включена</span><span class="sxs-lookup"><span data-stu-id="643d5-305">Enabled</span></span>|<span data-ttu-id="643d5-306">Отключена</span><span class="sxs-lookup"><span data-stu-id="643d5-306">Disabled</span></span>|<span data-ttu-id="643d5-307">Новый сеанс</span><span class="sxs-lookup"><span data-stu-id="643d5-307">New session</span></span>|<span data-ttu-id="643d5-308">Надстройки с функцией проверки при отправке не запускаются; отправка сообщения или элемента собрания.</span><span class="sxs-lookup"><span data-stu-id="643d5-308">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="643d5-309">Поддерживается.</span><span class="sxs-lookup"><span data-stu-id="643d5-309">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="643d5-310">Веб-браузер (современная версия Outlook), Windows, Mac</span><span class="sxs-lookup"><span data-stu-id="643d5-310">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="643d5-311">Чтобы внедрить функцию проверки при отправке, администраторы должны включить политику для обоих почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="643d5-311">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="643d5-312">Сведения о том, как поддерживать делегированный доступ в надстройке, см. в статье [Включение сценариев делегированного доступа в надстройке Outlook](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="643d5-312">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="643d5-313">Группа 1 — это современный почтовый ящик группы, а почтовый ящик пользователя 1 является участником группы 1</span><span class="sxs-lookup"><span data-stu-id="643d5-313">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="643d5-314">Сценарий</span><span class="sxs-lookup"><span data-stu-id="643d5-314">Scenario</span></span>|<span data-ttu-id="643d5-315">Политика проверки при отправке для почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="643d5-315">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="643d5-316">Включены ли надстройки, поддерживающие проверку сообщений при отправке?</span><span class="sxs-lookup"><span data-stu-id="643d5-316">On-send add-ins enabled?</span></span>|<span data-ttu-id="643d5-317">Действие почтового ящика 1</span><span class="sxs-lookup"><span data-stu-id="643d5-317">Mailbox 1 action</span></span>|<span data-ttu-id="643d5-318">Результат</span><span class="sxs-lookup"><span data-stu-id="643d5-318">Result</span></span>|<span data-ttu-id="643d5-319">Поддержка</span><span class="sxs-lookup"><span data-stu-id="643d5-319">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="643d5-320">1,1</span><span class="sxs-lookup"><span data-stu-id="643d5-320">1</span></span>|<span data-ttu-id="643d5-321">Включена</span><span class="sxs-lookup"><span data-stu-id="643d5-321">Enabled</span></span>|<span data-ttu-id="643d5-322">Да</span><span class="sxs-lookup"><span data-stu-id="643d5-322">Yes</span></span>|<span data-ttu-id="643d5-323">Почтовый ящик 1 создает новое сообщение или собрание для группы 1.</span><span class="sxs-lookup"><span data-stu-id="643d5-323">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="643d5-324">В случае отправки запускаются надстройки, поддерживающие проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-324">On-send add-ins run during send.</span></span>|<span data-ttu-id="643d5-325">Да</span><span class="sxs-lookup"><span data-stu-id="643d5-325">Yes</span></span>|
|<span data-ttu-id="643d5-326">2</span><span class="sxs-lookup"><span data-stu-id="643d5-326">2</span></span>|<span data-ttu-id="643d5-327">Включена</span><span class="sxs-lookup"><span data-stu-id="643d5-327">Enabled</span></span>|<span data-ttu-id="643d5-328">Да</span><span class="sxs-lookup"><span data-stu-id="643d5-328">Yes</span></span>|<span data-ttu-id="643d5-329">Почтовый ящик 1 создает новое сообщение или собрание для группы 1 в окне этой группы в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="643d5-329">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="643d5-330">В случае отправки не запускаются надстройки, поддерживающие проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-330">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="643d5-331">В настоящее время не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="643d5-331">Not currently supported.</span></span> <span data-ttu-id="643d5-332">В качестве обходного решения используйте сценарий 1.</span><span class="sxs-lookup"><span data-stu-id="643d5-332">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="643d5-333">Включен почтовый ящик пользователя с функцией или политикой проверки при отправке, установлены и включены надстройки, поддерживающие эту функцию, а также включен автономный режим</span><span class="sxs-lookup"><span data-stu-id="643d5-333">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="643d5-334">Надстройки, поддерживающие проверку при отправке, запускаются в соответствии с сетевым состоянием пользователя, внутреннего сервера надстройки и Exchange.</span><span class="sxs-lookup"><span data-stu-id="643d5-334">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="643d5-335">Состояние пользователя</span><span class="sxs-lookup"><span data-stu-id="643d5-335">User's state</span></span>

<span data-ttu-id="643d5-336">Надстройки, поддерживающие проверку сообщений при отправке, будут запускаться при отправке, если пользователь в сети.</span><span class="sxs-lookup"><span data-stu-id="643d5-336">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="643d5-337">В автономном режиме такие надстройки не будут запускаться при отправке, а сообщение или элемент собрания не будет отправлен.</span><span class="sxs-lookup"><span data-stu-id="643d5-337">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="643d5-338">Состояние внутреннего сервера надстройки</span><span class="sxs-lookup"><span data-stu-id="643d5-338">Add-in backend's state</span></span>

<span data-ttu-id="643d5-339">Надстройка, поддерживающая проверку при отправке, будет запускаться, если ее внутренний сервер подключен к сети и доступен.</span><span class="sxs-lookup"><span data-stu-id="643d5-339">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="643d5-340">Если внутренний сервер находится в автономном режиме, отправка отключена.</span><span class="sxs-lookup"><span data-stu-id="643d5-340">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="643d5-341">Состояние Exchange</span><span class="sxs-lookup"><span data-stu-id="643d5-341">Exchange's state</span></span>

<span data-ttu-id="643d5-342">Надстройки, поддерживающие проверку сообщений при отправке, будут запускаться при отправке, если сервер Exchange подключен к сети и доступен.</span><span class="sxs-lookup"><span data-stu-id="643d5-342">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="643d5-343">Если надстройке с функцией проверки при отправке недоступна служба Exchange и включена соответствующая политика или командлет, отправка отключена.</span><span class="sxs-lookup"><span data-stu-id="643d5-343">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="643d5-344">На компьютерах Mac в любом автономном состоянии кнопка **Отправить** (или **Отправить обновление** для существующих собраний) отключена, и отображается уведомление, что в организации не разрешено отправлять сообщения, если пользователь не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="643d5-344">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>


## <a name="code-examples"></a><span data-ttu-id="643d5-345">Примеры кода</span><span class="sxs-lookup"><span data-stu-id="643d5-345">Code examples</span></span>

<span data-ttu-id="643d5-346">В приведенных ниже примерах кода показано, как создать простую надстройку, поддерживающую проверку сообщений при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-346">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="643d5-347">Скачать код, на котором основаны эти примеры, можно на странице [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span><span class="sxs-lookup"><span data-stu-id="643d5-347">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="643d5-348">Манифест, переопределение версии и событие</span><span class="sxs-lookup"><span data-stu-id="643d5-348">Manifest, version override, and event</span></span>

<span data-ttu-id="643d5-349">Пример кода [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) включает два манифеста:</span><span class="sxs-lookup"><span data-stu-id="643d5-349">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="643d5-350">`Contoso Message Body Checker.xml` показывает, как проверить текст сообщения на наличие запрещенных слов или конфиденциальной информации при отправке.</span><span class="sxs-lookup"><span data-stu-id="643d5-350">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="643d5-351">`Contoso Subject and CC Checker.xml` показывает, как при отправке добавить получателя в строку "Копия" и проверить, включает ли сообщение строку темы.</span><span class="sxs-lookup"><span data-stu-id="643d5-351">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="643d5-352">В файле манифеста `Contoso Message Body Checker.xml` указываются файл и имя функции, которую следует вызывать при возникновении события `ItemSend`.</span><span class="sxs-lookup"><span data-stu-id="643d5-352">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="643d5-353">Операция выполняется синхронно.</span><span class="sxs-lookup"><span data-stu-id="643d5-353">The operation runs synchronously.</span></span>

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
> <span data-ttu-id="643d5-354">Если вы используете Visual Studio 2019 для разработки надстройки, используемой для отправки, вы можете получить предупреждение проверки следующего вида: "Это недопустимый xsi: Type"http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events. Чтобы обойти эту проблему, вам потребуется более новая версия MailAppVersionOverridesV1_1. xsd, которая была указана в [блоге в блоге](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/)GitHub.</span><span class="sxs-lookup"><span data-stu-id="643d5-354">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="643d5-355">Для файла манифеста `Contoso Subject and CC Checker.xml` в приведенном ниже примере показаны файл и имя функции, вызываемой при возникновении события отправки.</span><span class="sxs-lookup"><span data-stu-id="643d5-355">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

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

<span data-ttu-id="643d5-356">Для API проверки при отправке требуется узел `VersionOverrides v1_1`.</span><span class="sxs-lookup"><span data-stu-id="643d5-356">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="643d5-357">Ниже показано, как добавить узел `VersionOverrides` в манифест.</span><span class="sxs-lookup"><span data-stu-id="643d5-357">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On Send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="643d5-358">Дополнительные сведения см. в указанных ниже статьях.</span><span class="sxs-lookup"><span data-stu-id="643d5-358">For more information, see the following:</span></span>
> - [<span data-ttu-id="643d5-359">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="643d5-359">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="643d5-360">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="643d5-360">VersionOverrides</span></span>](../develop/create-addin-commands.md#step-3-add-versionoverrides-element)
> - [<span data-ttu-id="643d5-361">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="643d5-361">Office Add-ins XML manifest</span></span>](../overview/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="643d5-362">Объекты `Event` и `item`, методы `body.getAsync` и `body.setAsync`</span><span class="sxs-lookup"><span data-stu-id="643d5-362">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="643d5-363">Чтобы получить доступ к выбранному в данный момент сообщению или элементу собрания (в этом примере — к новому сообщению), используйте пространство имен `Office.context.mailbox.item`.</span><span class="sxs-lookup"><span data-stu-id="643d5-363">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="643d5-364">Функция проверки при отправке автоматически передает событие `ItemSend` функции, указанной в манифесте (в данном случае это функция `validateBody`).</span><span class="sxs-lookup"><span data-stu-id="643d5-364">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

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

<span data-ttu-id="643d5-365">Функция `validateBody` возвращает текущий текст в заданном формате (HTML) и передает нужный объект события `ItemSend` в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="643d5-365">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="643d5-366">Помимо метода `getAsync`, объект `Body` также предоставляет метод `setAsync`, с помощью которого вы можете заменить текст сообщения на указанный.</span><span class="sxs-lookup"><span data-stu-id="643d5-366">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="643d5-367">Дополнительные сведения см. в статьях [Объект Event](/javascript/api/office/office.addincommands.event) и [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="643d5-367">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="643d5-368">Объект `NotificationMessages` и метод `event.completed`</span><span class="sxs-lookup"><span data-stu-id="643d5-368">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="643d5-369">Функция `checkBodyOnlyOnSendCallBack` использует регулярное выражение, чтобы определить, содержит ли текст сообщения слова, подлежащие блокировке.</span><span class="sxs-lookup"><span data-stu-id="643d5-369">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="643d5-370">Если она обнаруживает слово, совпадающие с каким-либо элементом из массива запрещенных слов, отправка сообщения блокируется, а отправитель получает уведомление на панели информации.</span><span class="sxs-lookup"><span data-stu-id="643d5-370">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="643d5-371">Для этого в ней используется свойство `notificationMessages` объекта `Item` для возврата объекта `NotificationMessages`.</span><span class="sxs-lookup"><span data-stu-id="643d5-371">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="643d5-372">После этого она добавляет уведомление к элементу, вызывая метод `addAsync`, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="643d5-372">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

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

<span data-ttu-id="643d5-373">Ниже перечислены параметры метода `addAsync`.</span><span class="sxs-lookup"><span data-stu-id="643d5-373">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="643d5-374">`NoSend`. Строка, представляющая собой заданный разработчиком ключ для ссылки на сообщение уведомления.</span><span class="sxs-lookup"><span data-stu-id="643d5-374">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="643d5-375">С его помощью вы сможете изменить это сообщение позже.</span><span class="sxs-lookup"><span data-stu-id="643d5-375">You can use it to modify this message later.</span></span> <span data-ttu-id="643d5-376">Ключ не может быть длиннее 32 символов.</span><span class="sxs-lookup"><span data-stu-id="643d5-376">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="643d5-377">`type`. Одно из свойств параметра объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="643d5-377">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="643d5-378">Представляет тип сообщения. Типы соответствуют значениям перечисления [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype).</span><span class="sxs-lookup"><span data-stu-id="643d5-378">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="643d5-379">Допустимые значения: индикатор хода выполнения, информационное сообщение и сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="643d5-379">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="643d5-380">В этом примере в свойстве `type` указано сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="643d5-380">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="643d5-381">`message`. Одно из свойств параметра объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="643d5-381">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="643d5-382">В этом примере `message` — это текст сообщения уведомления.</span><span class="sxs-lookup"><span data-stu-id="643d5-382">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="643d5-383">Чтобы сообщить о завершении надстройкой обработки события `ItemSend`, активированного операцией отправки, вызовите метод `event.completed({allowEvent:Boolean})`.</span><span class="sxs-lookup"><span data-stu-id="643d5-383">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="643d5-384">Свойство `allowEvent` является логическим.</span><span class="sxs-lookup"><span data-stu-id="643d5-384">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="643d5-385">Если задано значение `true`, отправка разрешается.</span><span class="sxs-lookup"><span data-stu-id="643d5-385">If set to `true`, send is allowed.</span></span> <span data-ttu-id="643d5-386">Если задано значение `false`, отправка письма блокируется.</span><span class="sxs-lookup"><span data-stu-id="643d5-386">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="643d5-387">Дополнительные сведения см. в статьях [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [completed](/javascript/api/office/office.addincommands.event).</span><span class="sxs-lookup"><span data-stu-id="643d5-387">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="643d5-388">Методы `replaceAsync`, `removeAsync` и `getAllAsync`</span><span class="sxs-lookup"><span data-stu-id="643d5-388">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="643d5-389">Помимо метода `addAsync`, объект `NotificationMessages` также включает методы `replaceAsync`, `removeAsync` и `getAllAsync`.</span><span class="sxs-lookup"><span data-stu-id="643d5-389">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="643d5-390">Эти методы не используются в данном примере кода.</span><span class="sxs-lookup"><span data-stu-id="643d5-390">These methods are not used in this code sample.</span></span>  <span data-ttu-id="643d5-391">Дополнительные сведения см. в статье [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span><span class="sxs-lookup"><span data-stu-id="643d5-391">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="643d5-392">Код проверки строк "Тема" и "Копия"</span><span class="sxs-lookup"><span data-stu-id="643d5-392">Subject and CC checker code</span></span>

<span data-ttu-id="643d5-393">В приведенном ниже примере кода показано, как при отправке сообщения добавить получателя в строку "Копия" и проверить, включает ли сообщение тему.</span><span class="sxs-lookup"><span data-stu-id="643d5-393">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="643d5-394">В этом примере функция проверки при отправке используется, чтобы разрешить или запретить отправку сообщения.</span><span class="sxs-lookup"><span data-stu-id="643d5-394">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

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

<span data-ttu-id="643d5-p152">Дополнительные сведения о том, как при отправке сообщения добавить получателя в строку "Копия" и проверить, указана ли тема сообщения, а также просмотреть доступные API, см. в [примере Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). Код сопровождается подробными комментариями.</span><span class="sxs-lookup"><span data-stu-id="643d5-p152">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="643d5-397">См. также</span><span class="sxs-lookup"><span data-stu-id="643d5-397">See also</span></span>

- [<span data-ttu-id="643d5-398">Обзор архитектуры и функций надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="643d5-398">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="643d5-399">Надстройка Outlook "Демонстрация команд надстройки"</span><span class="sxs-lookup"><span data-stu-id="643d5-399">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)
