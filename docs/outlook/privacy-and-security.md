---
title: Конфиденциальность, разрешения и безопасность для надстроек Outlook
description: Узнайте, как управлять конфиденциальностью, разрешениями и безопасностью в надстройке Outlook.
ms.date: 10/31/2019
localization_priority: Priority
ms.openlocfilehash: d233eb3ac6980af24e6ba9d951834532ea79dc06
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605334"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a><span data-ttu-id="83837-103">Конфиденциальность, разрешения и безопасность для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="83837-103">Privacy, permissions, and security for Outlook add-ins</span></span>

<span data-ttu-id="83837-104">Пользователи, разработчики и администраторы могут использовать уровни разрешений модели безопасности для надстроек Outlook, чтобы управлять конфиденциальностью и производительностью.</span><span class="sxs-lookup"><span data-stu-id="83837-104">End users, developers, and administrators can use the tiered permission levels of the security model for Outlook add-ins to control privacy and performance.</span></span>

<span data-ttu-id="83837-105">В этой статье описаны разрешения, которые могут запрашивать надстройки Outlook, и рассматривается модель безопасности с различных точек зрения:</span><span class="sxs-lookup"><span data-stu-id="83837-105">This article describes the possible permissions that Outlook add-ins can request, and examines the security model from the following perspectives:</span></span>

- <span data-ttu-id="83837-106">**AppSource**: целостность надстройки</span><span class="sxs-lookup"><span data-stu-id="83837-106">**AppSource**: add-in integrity</span></span>
    
- <span data-ttu-id="83837-107">**Пользователи**: вопросы, связанные с конфиденциальностью и производительностью</span><span class="sxs-lookup"><span data-stu-id="83837-107">**End-users**: privacy and performance concerns</span></span>
    
- <span data-ttu-id="83837-108">**Разработчики**: варианты разрешений и ограничения на использование ресурсов</span><span class="sxs-lookup"><span data-stu-id="83837-108">**Developers**: permissions choices and resource usage limits</span></span>
    
- <span data-ttu-id="83837-109">**Администраторы**: разрешения на определение пороговых значений производительности</span><span class="sxs-lookup"><span data-stu-id="83837-109">**Administrators**: privileges to set performance thresholds</span></span>
    

## <a name="permissions-model"></a><span data-ttu-id="83837-110">Модель разрешений</span><span class="sxs-lookup"><span data-stu-id="83837-110">Permissions model</span></span>

<span data-ttu-id="83837-p101">От того, насколько клиенты доверяют безопасности надстройки, зависит ее принятие. Безопасность надстройки Outlook опирается на уровневую модель разрешений. Надстройка Outlook открывает необходимый уровень разрешений, определяя возможный доступ и действия, которые она может выполнять с данными почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="83837-p101">Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.</span></span> 

<span data-ttu-id="83837-113">Схема манифеста версии 1.1 включает четыре уровня разрешений.</span><span class="sxs-lookup"><span data-stu-id="83837-113">Manifest schema version 1.1 includes four levels of permissions.</span></span> 


<span data-ttu-id="83837-114">**Таблица 1. Уровни разрешений для надстройки**</span><span class="sxs-lookup"><span data-stu-id="83837-114">**Table 1. Add-in permission levels**</span></span>

|<span data-ttu-id="83837-115">**Уровень разрешений**</span><span class="sxs-lookup"><span data-stu-id="83837-115">**Permission level**</span></span>|<span data-ttu-id="83837-116">**Значение в манифесте надстройки Outlook**</span><span class="sxs-lookup"><span data-stu-id="83837-116">**Value in Outlook add-in manifest**</span></span>|
|:-----|:-----|
|<span data-ttu-id="83837-117">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="83837-117">Restricted</span></span>|<span data-ttu-id="83837-118">Restricted</span><span class="sxs-lookup"><span data-stu-id="83837-118">Restricted</span></span>|
|<span data-ttu-id="83837-119">Чтение элемента</span><span class="sxs-lookup"><span data-stu-id="83837-119">Read item</span></span>|<span data-ttu-id="83837-120">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83837-120">ReadItem</span></span>|
|<span data-ttu-id="83837-121">Чтение и запись элемента</span><span class="sxs-lookup"><span data-stu-id="83837-121">Read/write item</span></span>|<span data-ttu-id="83837-122">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="83837-122">ReadWriteItem</span></span>|
|<span data-ttu-id="83837-123">Чтение и запись почтового ящика</span><span class="sxs-lookup"><span data-stu-id="83837-123">Read/write mailbox</span></span>|<span data-ttu-id="83837-124">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="83837-124">ReadWriteMailbox</span></span>|

<span data-ttu-id="83837-125">Четыре уровня разрешений являются накопительными: разрешение **на чтение и запись в почтовом ящике** включает в себя разрешения **на чтение и запись элемента**, **на чтение элемента** и **на ограниченный доступ**; разрешение **на чтение и запись элемента** включает в себя разрешения **на чтение элемента** и **на ограниченный доступ**; разрешение **на чтение элемента** включает в себя разрешение **на ограниченный доступ**.</span><span class="sxs-lookup"><span data-stu-id="83837-125">The four levels of permissions are cumulative: the **read/write mailbox** permission includes the permissions of **read/write item**, **read item** and **restricted**, **read/write item** includes **read item** and **restricted**, and the **read item** permission includes **restricted**.</span></span> 

<span data-ttu-id="83837-126">На рисунке ниже показано четыре уровня разрешений и описаны возможности, обеспечиваемые каждым уровнем для пользователей, разработчиков и администраторов.</span><span class="sxs-lookup"><span data-stu-id="83837-126">The following figure shows the four levels of permissions and describes the capabilities offered to the end user, developer, and administrator by each tier.</span></span> <span data-ttu-id="83837-127">Дополнительные сведения об этих разрешениях см. в статьях [Пользователи: вопросы, связанные с конфиденциальностью и производительностью](#end-users-privacy-and-performance-concerns), [Разработчики: варианты разрешений и ограничения на использование ресурсов](#developers-permission-choices-and-resource-usage-limits) и [Общие сведения о разрешениях для надстроек Outlook](understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="83837-127">For more information about these permissions, see [End users: privacy and performance concerns](#end-users-privacy-and-performance-concerns), [Developers: permission choices and resource usage limits](#developers-permission-choices-and-resource-usage-limits), and [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span> 


<span data-ttu-id="83837-128">**Сопоставление четырехуровневой модели разрешений с пользователями, разработчиками и администраторами**</span><span class="sxs-lookup"><span data-stu-id="83837-128">**Relating the four-tier permission model to the end user, developer, and administrator**</span></span>

![Четырехуровневая модель разрешений для схемы почтовых приложений версии 1.1](../images/add-in-permission-tiers.png)


## <a name="appsource-add-in-integrity"></a><span data-ttu-id="83837-130">AppSource: целостность надстройки</span><span class="sxs-lookup"><span data-stu-id="83837-130">AppSource: add-in integrity</span></span>

<span data-ttu-id="83837-131">[AppSource](https://appsource.microsoft.com) содержит надстройки, которые могут установить пользователи и администраторы.</span><span class="sxs-lookup"><span data-stu-id="83837-131">[AppSource](https://appsource.microsoft.com) hosts add-ins that can be installed by end users and administrators.</span></span> <span data-ttu-id="83837-132">AppSource применяет указанные ниже меры для поддержки целостности надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="83837-132">AppSource enforces the following measures to maintain the integrity of these Outlook add-ins:</span></span>

- <span data-ttu-id="83837-133">Требуется постоянное использование сервером, на котором размещена надстройка, протокола SSL для связи.</span><span class="sxs-lookup"><span data-stu-id="83837-133">Requires the host server of an add-in to always use Secure Socket Layer (SSL) to communicate.</span></span>
    
- <span data-ttu-id="83837-134">Разработчику требуется предоставить доказательство подлинности, контрактное соглашение и соответствующую требованиям политику конфиденциальности для отправки надстроек.</span><span class="sxs-lookup"><span data-stu-id="83837-134">Requires a developer to provide proof of identity, a contractual agreement, and a compliant privacy policy to submit add-ins.</span></span> 
    
- <span data-ttu-id="83837-135">Архивация надстроек осуществляется только в режиме для чтения.</span><span class="sxs-lookup"><span data-stu-id="83837-135">Archives add-ins in read-only mode.</span></span>
    
- <span data-ttu-id="83837-136">Поддерживается система рецензий пользователей для доступных надстроек, чтобы стимулировать формирование саморегулирующегося сообщества.</span><span class="sxs-lookup"><span data-stu-id="83837-136">Supports a user-review system for available add-ins to promote a self-policing community.</span></span>
    

## <a name="end-users-privacy-and-performance-concerns"></a><span data-ttu-id="83837-137">Пользователи: вопросы, связанные с конфиденциальностью и производительностью</span><span class="sxs-lookup"><span data-stu-id="83837-137">End users: privacy and performance concerns</span></span>

<span data-ttu-id="83837-138">Модель безопасности адресует вопросы пользователей, связанные с безопасностью, конфиденциальностью и производительностью, следующим образом.</span><span class="sxs-lookup"><span data-stu-id="83837-138">The security model addresses security, privacy, and performance concerns of end users in the following ways:</span></span>

- <span data-ttu-id="83837-139">Пользовательские сообщения, защищенные с помощью IRM в Outlook, не взаимодействуют с надстройками Outlook.</span><span class="sxs-lookup"><span data-stu-id="83837-139">End user's messages that are protected by Outlook's Information Rights Management (IRM) do not interact with Outlook add-ins.</span></span>
    
- <span data-ttu-id="83837-140">Перед установкой надстройки, добавленной в AppSource, пользователям видны сведения о доступе и действиях, которые надстройка может выполнять с их данными, и для продолжения установки необходимо явно подтвердить свое согласие.</span><span class="sxs-lookup"><span data-stu-id="83837-140">Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed.</span></span> <span data-ttu-id="83837-141">Ни одна надстройка Outlook не устанавливается на клиентский компьютер без получения непосредственного согласия от пользователя или администратора.</span><span class="sxs-lookup"><span data-stu-id="83837-141">No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.</span></span>
    
- <span data-ttu-id="83837-p105">Разрешение **ограниченное** позволяет ограничить доступ надстройки Outlook только к текущему элементу. Разрешение **чтение элемента** позволяет надстройке получить доступ к личным сведениям, например именам и электронным адресам отправителя и получателя, но только для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="83837-p105">Granting the **restricted** permission allows the Outlook add-in to have limited access on only the current item. Granting the **read item** permission allows the Outlook add-in to access personal identifiable information, such as sender and recipient names and email addresses, on only the current item,.</span></span>
    
- <span data-ttu-id="83837-p106">Пользователь может установить надстройку Outlook только для себя. Установку надстроек Outlook на уровне всей организации выполняют администраторы.</span><span class="sxs-lookup"><span data-stu-id="83837-p106">An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.</span></span>
    
- <span data-ttu-id="83837-146">Пользователи могут устанавливать надстройки Outlook, которые задействуют сценарии, зависимые от контекста, что очень привлекательно для самих пользователей, но при этом также снижает риски, связанные с безопасностью.</span><span class="sxs-lookup"><span data-stu-id="83837-146">End users can install Outlook add-ins that enable context-sensitive scenarios that are compelling to users while minimizing the users' security risks.</span></span>
    
- <span data-ttu-id="83837-147">Защита файлов манифестов установленных надстроек Outlook обеспечивается в учетной записи электронной почты пользователя.</span><span class="sxs-lookup"><span data-stu-id="83837-147">Manifest files of installed Outlook add-ins are secured in the user's email account.</span></span>
    
- <span data-ttu-id="83837-148">Данные, которыми приложения обмениваются с серверами, на которых установлены Надстройки Office, всегда шифруются в соответствии с протоколом SSL.</span><span class="sxs-lookup"><span data-stu-id="83837-148">Data communicated with servers hosting Office Add-ins is always encrypted according to the Secure Socket Layer (SSL) protocol.</span></span>
    
- <span data-ttu-id="83837-149">Применимо только к полнофункциональным клиентам Outlook, которые отслеживают производительность установленных надстроек Outlook, контролируют их и отключают те приложения, которые превышают ограничения по ряду следующих факторов:</span><span class="sxs-lookup"><span data-stu-id="83837-149">Applicable to only the Outlook rich clients: The Outlook rich clients monitor the performance of installed Outlook add-ins, exercise governance control, and disable those Outlook add-ins that exceed limits in the following areas:</span></span>
    
  - <span data-ttu-id="83837-150">Время отзыва для активации</span><span class="sxs-lookup"><span data-stu-id="83837-150">Response time to activate</span></span>
    
  - <span data-ttu-id="83837-151">Количество сбоев при активации или повторной активации</span><span class="sxs-lookup"><span data-stu-id="83837-151">Number of failures to activate or reactivate</span></span>
    
  - <span data-ttu-id="83837-152">Использование памяти</span><span class="sxs-lookup"><span data-stu-id="83837-152">Memory usage</span></span>
    
  - <span data-ttu-id="83837-153">Использование процессора</span><span class="sxs-lookup"><span data-stu-id="83837-153">CPU usage</span></span>  

  <span data-ttu-id="83837-p107">Такой контроль предотвращает атаки по принципу отказа в обслуживании и поддерживает производительность надстройки на допустимом уровне. В бизнес-строке пользователи получают уведомления о тех надстройках Outlook, которые полнофункциональный клиент Outlook отключил, руководствуясь изложенными выше принципами.</span><span class="sxs-lookup"><span data-stu-id="83837-p107">Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.</span></span>

- <span data-ttu-id="83837-156">В любой момент пользователи могут проверить разрешения, запрашиваемые установленными надстройками Outlook, отключить, а затем включить любую надстройку Outlook в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="83837-156">At any time, end users can verify the permissions requested by installed Outlook add-ins, and disable or subsequently enable any Outlook add-in in the Exchange Admin Center.</span></span>


## <a name="developers-permission-choices-and-resource-usage-limits"></a><span data-ttu-id="83837-157">Разработчики — варианты разрешений и ограничения на использование ресурсов</span><span class="sxs-lookup"><span data-stu-id="83837-157">Developers: permission choices and resource usage limits</span></span>

<span data-ttu-id="83837-158">Модель безопасности предоставляет пользователям возможность детальной настройки разрешений, а также строгие правила производительности, которых следует придерживаться.</span><span class="sxs-lookup"><span data-stu-id="83837-158">The security model provides developers granular levels of permissions to choose from, and strict performance guidelines to observe.</span></span>

### <a name="tiered-permissions-increases-transparency"></a><span data-ttu-id="83837-159">Уровневая модель разрешений повышает прозрачность</span><span class="sxs-lookup"><span data-stu-id="83837-159">Tiered permissions increases transparency</span></span>

<span data-ttu-id="83837-160">Разработчикам рекомендуется следовать многоуровневой модели разрешений, чтобы обеспечить прозрачность и развеять сомнения пользователей по поводу доступа надстроек к их данным и почтовым ящикам, что косвенно повысит популярность надстройки:</span><span class="sxs-lookup"><span data-stu-id="83837-160">Developers should follow the tiered permissions model to provide transparency and alleviate users' concern about what add-ins can do to their data and mailbox, indirectly promoting add-in adoption:</span></span>

- <span data-ttu-id="83837-161">Разработчики запрашивают подходящий уровень разрешений для надстройки Outlook с учетом способа ее активации, а также необходимости чтения или записи определенных свойств элемента или создания и отправки элемента.</span><span class="sxs-lookup"><span data-stu-id="83837-161">Developers request an appropriate level of permission for an Outlook add-in, based on how the Outlook add-in should be activated, and its need to read or write certain properties of an item, or to create and send an item.</span></span>

- <span data-ttu-id="83837-162">Разработчики запрашивают разрешение с помощью элемента [Permissions](../reference/manifest/permissions.md) в манифесте надстройки Outlook, назначая значение **Restricted**, **ReadItem**, **ReadWriteItem** или **ReadWriteMailbox**.</span><span class="sxs-lookup"><span data-stu-id="83837-162">Developers request permission by using the [Permissions](../reference/manifest/permissions.md) element in the manifest of the Outlook add-in, by assigning a value of **Restricted**, **ReadItem**, **ReadWriteItem** or **ReadWriteMailbox**, as appropriate.</span></span>

  > [!NOTE]
  > <span data-ttu-id="83837-163">Помните, что разрешение **ReadWriteItem** доступно начиная со схемы манифеста версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="83837-163">Note that the **ReadWriteItem** permission is available starting in manifest schema v1.1.</span></span>

  <span data-ttu-id="83837-164">В приведенном ниже примере запрашивается разрешение на **чтение элемента**.</span><span class="sxs-lookup"><span data-stu-id="83837-164">The following example requests the **read item** permission.</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- <span data-ttu-id="83837-165">Разработчики могут запрашивать разрешение **на ограниченный доступ**, если надстройка Outlook задействуется только для определенного типа элементов Outlook (встреча или сообщение) или для определенных извлеченных сущностей (телефонный номер, адрес, URL-адрес) в теме или основном тексте элемента.</span><span class="sxs-lookup"><span data-stu-id="83837-165">Developers can request the **restricted** permission if the Outlook add-in activates on a specific type of Outlook items (appointment or message), or on specific extracted entities (phone number, address, URL) being present in the item's subject or body.</span></span> <span data-ttu-id="83837-166">Например, указанное ниже правило активирует надстройку Outlook, если хотя бы одна из трех сущностей (телефонный номер, почтовый адрес или URL-адрес) найдена в теме или основном тексте текущего сообщения.</span><span class="sxs-lookup"><span data-stu-id="83837-166">For example, the following rule activates the Outlook add-in if one or more of three entities - phone number, postal address, or URL - are found in the subject or body of the current message.</span></span>
    
  ```XML
    <Permissions>Restricted</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        </Rule>
    </Rule>
  ```

- <span data-ttu-id="83837-167">Разработчикам следует запрашивать разрешение на **чтение элемента**, если надстройка Outlook должна считывать свойства текущего элемента, который не входит в извлеченные сущности по умолчанию, или записывать пользовательские свойства, определяемые надстройкой для текущего элемента, но не требует чтения или записи других элементов либо создания и отправки сообщения в пользовательском почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="83837-167">Developers should request the **read item** permission if the Outlook add-in needs to read properties of the current item other than the default extracted entities, or write custom properties set by the add-in on the current item, but does not require reading or writing to other items, or creating or sending a message in the user's mailbox.</span></span> <span data-ttu-id="83837-168">Например, разработчик должен запросить разрешение на **чтение элемента**, если надстройка должна искать такие сущности, как приглашение на собрание, предложение задачи, электронный адрес или имя контакта в теме или основном тексте элемента, или если для активации надстройки требуется регулярное выражение.</span><span class="sxs-lookup"><span data-stu-id="83837-168">For example, a developer should request **read item** permission if an Outlook add-in needs to look for an entity like a meeting suggestion, task suggestion, email address, or contact name in the item's subject or body, or uses a regular expression to activate.</span></span>

- <span data-ttu-id="83837-169">Разработчикам следует запрашивать разрешение на **чтение и запись элемента**, если надстройка Outlook должна записывать свойства созданного элемента, например имена, электронные адреса, основной текст и тему или добавлять и удалять вложения.</span><span class="sxs-lookup"><span data-stu-id="83837-169">Developers should request the **read/write item** permission if the Outlook add-in needs to write to properties of the composed item, such as recipient names, email addresses, body, and subject, or needs to add or remove item attachments.</span></span>

- <span data-ttu-id="83837-170">Разработчики запрашивают разрешение **чтение и запись в почтовом ящике**, только если надстройка Outlook должна выполнять одно или несколько из приведенных ниже действий с помощью метода [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="83837-170">Developers request the **read/write mailbox** permission only if the Outlook add-in needs to do one or more of the following actions by using the [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method:</span></span>

  - <span data-ttu-id="83837-171">Чтение и запись свойств элементов в почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="83837-171">Read or write to properties of items in the mailbox.</span></span>
  - <span data-ttu-id="83837-172">Создание, чтение, запись и отправка элементов в почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="83837-172">Create, read, write, or send items in the mailbox.</span></span>
  - <span data-ttu-id="83837-173">Создание, чтение папок почтового ящика и запись в них.</span><span class="sxs-lookup"><span data-stu-id="83837-173">Create, read, or write to folders in the mailbox.</span></span>


### <a name="resource-usage-tuning"></a><span data-ttu-id="83837-174">Регулирование использования ресурсов</span><span class="sxs-lookup"><span data-stu-id="83837-174">Resource usage tuning</span></span>

<span data-ttu-id="83837-p110">Разработчики должны знать пределы использования ресурсов для активации и учитывать необходимость оптимальной настройки производительности в рабочем процессе разработки, чтобы снизить вероятность отказа в обслуживании из-за низкой производительности надстройки. Рекомендуем следовать инструкциям по разработке правил активации, представленным в статье [Ограничения для активации и API JavaScript для надстроек Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). Если надстройка Outlook должна работать в полнофункциональном клиенте Outlook, разработчикам стоит убедиться, что она правильно работает при ограниченном использовании ресурсов.</span><span class="sxs-lookup"><span data-stu-id="83837-p110">Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.</span></span>


### <a name="other-measures-to-promote-user-security"></a><span data-ttu-id="83837-177">Другие меры повышения безопасности пользователей</span><span class="sxs-lookup"><span data-stu-id="83837-177">Other measures to promote user security</span></span>

<span data-ttu-id="83837-178">Разработчики также должны знать и учитывать следующее:</span><span class="sxs-lookup"><span data-stu-id="83837-178">Developers should be aware of and plan for the following as well:</span></span>

- <span data-ttu-id="83837-179">Разработчики не могут использовать элементы ActiveX в своих надстройках, так как эти элементы не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="83837-179">Developers cannot use ActiveX controls in add-ins because they are not supported.</span></span>
    
- <span data-ttu-id="83837-180">При отправке надстройки Outlook в AppSource разработчикам следует:</span><span class="sxs-lookup"><span data-stu-id="83837-180">Developers should do the following when submitting an Outlook add-in to AppSource:</span></span>
    
  - <span data-ttu-id="83837-181">Создать SSL-сертификат высокой надежности для подтверждения своей личности.</span><span class="sxs-lookup"><span data-stu-id="83837-181">Produce an Extended Validation (EV) SSL certificate as a proof of identity.</span></span>
    
  - <span data-ttu-id="83837-182">Разместить предоставляемую надстройку на веб-сервере, поддерживающем SSL.</span><span class="sxs-lookup"><span data-stu-id="83837-182">Host the add-in they are submitting on a web server that supports SSL.</span></span>
    
  - <span data-ttu-id="83837-183">Создать соответствующую нормам политику конфиденциальности.</span><span class="sxs-lookup"><span data-stu-id="83837-183">Produce a compliant privacy policy.</span></span>
    
  - <span data-ttu-id="83837-184">Быть готовыми подписать договорное соглашение при предоставлении надстройки.</span><span class="sxs-lookup"><span data-stu-id="83837-184">Be ready to sign a contractual agreement upon submitting the add-in.</span></span>
    

## <a name="administrators-privileges"></a><span data-ttu-id="83837-185">Администраторы: привилегии</span><span class="sxs-lookup"><span data-stu-id="83837-185">Administrators: privileges</span></span>

<span data-ttu-id="83837-186">Модель разработки предоставляет администраторам следующие права и обязанности.</span><span class="sxs-lookup"><span data-stu-id="83837-186">The security model provides the following rights and responsibilities to administrators:</span></span>

- <span data-ttu-id="83837-187">Возможность запретить пользователям устанавливать какие-либо надстройки Outlook, в том числе надстройки из AppSource.</span><span class="sxs-lookup"><span data-stu-id="83837-187">Can prevent end users from installing any Outlook add-in, including add-ins from AppSource.</span></span>
    
- <span data-ttu-id="83837-188">Возможность отключать или включать любую надстройку Outlook в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="83837-188">Can disable or enable any Outlook add-in on the Exchange Admin Center.</span></span>
    
- <span data-ttu-id="83837-189">Применимо только к Outlook для Windows: можно переопределить параметры пороговых значений производительности с помощью параметров реестра в объекте глобальной политики.</span><span class="sxs-lookup"><span data-stu-id="83837-189">Applicable to only Outlook on Windows: Can override performance threshold settings by GPO registry settings.</span></span>
    


## <a name="see-also"></a><span data-ttu-id="83837-190">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="83837-190">See also</span></span>

- [<span data-ttu-id="83837-191">Конфиденциальность и безопасность надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="83837-191">Privacy and security for Office Add-ins</span></span>](../develop/privacy-and-security.md)    
- [<span data-ttu-id="83837-192">API надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="83837-192">Outlook add-in APIs</span></span>](apis.md)    
- [<span data-ttu-id="83837-193">Ограничения для активации и API JavaScript для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="83837-193">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
