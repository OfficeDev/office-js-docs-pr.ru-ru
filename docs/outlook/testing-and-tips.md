---
title: Развертывание и установка надстроек Outlook для тестирования
description: Создайте файл манифеста, разверните файл пользовательского интерфейса надстройки на веб-сервере, установите надстройку в своем почтовом ящике, а затем протестируйте ее.
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 76688ad3e1eca2dda832a94c3a9ae815e37678bc
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890979"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a><span data-ttu-id="dc2be-103">Развертывание и установка надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="dc2be-103">Deploy and install Outlook add-ins for testing</span></span>

<span data-ttu-id="dc2be-104">В рамках разработки надстройки Outlook вам, скорее всего, понадобится несколько раз развертывать и устанавливать надстройку для тестирования, что подразумевает выполнение следующих действий.</span><span class="sxs-lookup"><span data-stu-id="dc2be-104">As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps:</span></span>

1. <span data-ttu-id="dc2be-105">Создание файла манифеста, в котором описывается надстройка.</span><span class="sxs-lookup"><span data-stu-id="dc2be-105">Creating a manifest file that describes the add-in.</span></span>
1. <span data-ttu-id="dc2be-106">Развертывание файлов пользовательского интерфейса надстройки на веб-сервере.</span><span class="sxs-lookup"><span data-stu-id="dc2be-106">Deploying the add-in UI file(s) to a web server.</span></span>
1. <span data-ttu-id="dc2be-107">Установка надстройки в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="dc2be-107">Installing the add-in in your mailbox.</span></span>
1. <span data-ttu-id="dc2be-108">Тестирование надстройки с внесением соответствующих изменений в пользовательский интерфейс или файлы манифеста и повторение этапов 2 и 3 для тестирования изменений.</span><span class="sxs-lookup"><span data-stu-id="dc2be-108">Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.</span></span>

> [!NOTE]
> <span data-ttu-id="dc2be-109">Поскольку [настраиваемые области устарели](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), следует убедиться, что вы используете [поддерживаемую точку расширения надстройки](outlook-add-ins-overview.md#extension-points).</span><span class="sxs-lookup"><span data-stu-id="dc2be-109">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).</span></span>

## <a name="create-a-manifest-file-for-the-add-in"></a><span data-ttu-id="dc2be-110">Создание файла манифеста для надстройки</span><span class="sxs-lookup"><span data-stu-id="dc2be-110">Create a manifest file for the add-in</span></span>

<span data-ttu-id="dc2be-p101">Каждая надстройка описывается XML-манифестом, то есть документом, который предоставляет серверу сведения о почтовой надстройке, сообщает пользователям подробные сведения о надстройке и определяет местоположение HTML-файла пользовательского интерфейса надстройки. Вы можете сохранить манифест в локальной папке или на сервере, если у сервера Exchange, где размещен почтовый ящик, используемый в тестировании, есть доступ к этому месту. Сведения о создании файла манифеста см. в разделе [Манифесты надстроек Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="dc2be-p101">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="deploy-an-add-in-to-a-web-server"></a><span data-ttu-id="dc2be-115">Развертывание надстройки на веб-сервере</span><span class="sxs-lookup"><span data-stu-id="dc2be-115">Deploy an add-in to a web server</span></span>

<span data-ttu-id="dc2be-p102">Для создания надстройки можно использовать HTML и JavaScript. Конечные исходные файлы хранятся на веб-сервере, к которому может обращаться сервер Exchange Server, на котором размещена надстройка. После развертывания исходных файлов надстройки вы можете обновить ее пользовательский интерфейс и поведение, обновив файлы HTML или JavaScript на веб-сервере.</span><span class="sxs-lookup"><span data-stu-id="dc2be-p102">You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span></span>

## <a name="install-the-add-in"></a><span data-ttu-id="dc2be-119">Установка надстройки</span><span class="sxs-lookup"><span data-stu-id="dc2be-119">Install the add-in</span></span>

<span data-ttu-id="dc2be-120">После подготовки файла манифеста и развертывания пользовательского интерфейса надстройки на доступном веб-сервере, вы можете загрузить неопубликованную надстройку для почтового ящика на сервере Exchange Server, используя клиент Outlook, или установить ее с помощью командлетов Windows PowerShell.</span><span class="sxs-lookup"><span data-stu-id="dc2be-120">After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.</span></span>

### <a name="sideload-the-add-in"></a><span data-ttu-id="dc2be-121">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="dc2be-121">Sideload the add-in</span></span>

<span data-ttu-id="dc2be-p103">Вы можете установить надстройку, если ваш почтовый ящик находится в Exchange Online, Exchange 2013 или более поздней версии. Для загрузки неопубликованных надстроек требуется по крайней мере роль **My Custom Apps** для сервера Exchange Server. Чтобы проверить надстройку или иметь возможность устанавливать надстройки, указывая URL-адрес или имя файла манифеста, попросите своего администратора Exchange предоставить вам необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="dc2be-p103">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span></span>

<span data-ttu-id="dc2be-p104">Администратор Exchange может выполнить следующий командлет PowerShell, чтобы назначить необходимые разрешения одному пользователю. В этом примере `wendyri` — псевдоним электронной почты пользователя.</span><span class="sxs-lookup"><span data-stu-id="dc2be-p104">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.</span></span>

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

<span data-ttu-id="dc2be-127">При необходимости администратор может выполнить следующий командлет, чтобы назначить похожие разрешения нескольким пользователям:</span><span class="sxs-lookup"><span data-stu-id="dc2be-127">If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions:</span></span>

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

<span data-ttu-id="dc2be-128">Дополнительные сведения об упомянутой роли см. в статье [Роль My Custom Apps](/exchange/my-custom-apps-role-exchange-2013-help).</span><span class="sxs-lookup"><span data-stu-id="dc2be-128">For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).</span></span>

<span data-ttu-id="dc2be-129">Если для разработки надстроек вы используете Office 365 или Visual Studio, вам назначается роль администратора организации, позволяющая устанавливать надстройки с помощью файла или URL-адреса в Центре администрирования Exchange, а также с помощью командлетов PowerShell.</span><span class="sxs-lookup"><span data-stu-id="dc2be-129">Using Office 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.</span></span>

### <a name="install-an-add-in-by-using-remote-powershell"></a><span data-ttu-id="dc2be-130">Установка надстройки с помощью удаленного сеанса PowerShell</span><span class="sxs-lookup"><span data-stu-id="dc2be-130">Install an add-in by using remote PowerShell</span></span>

<span data-ttu-id="dc2be-131">После создания удаленного сеанса Windows PowerShell на сервере Exchange Server вы можете установить надстройку Outlook, используя командлет `New-App` и следующую команду PowerShell.</span><span class="sxs-lookup"><span data-stu-id="dc2be-131">After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.</span></span>

```powershell
New-App -URL:"http://<fully-qualified URL">
```

<span data-ttu-id="dc2be-132">Полный URL-адрес — это расположение подготовленного файла манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="dc2be-132">The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.</span></span>

<span data-ttu-id="dc2be-133">Вы можете использовать следующие командлеты PowerShell для управления надстройками для почтового ящика:</span><span class="sxs-lookup"><span data-stu-id="dc2be-133">You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox:</span></span>

-  <span data-ttu-id="dc2be-134">`Get-App`: отображает надстройки, включенные для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="dc2be-134">`Get-App` - Lists the add-ins that are enabled for a mailbox.</span></span>
-  <span data-ttu-id="dc2be-135">`Set-App`: включает или отключает надстройку для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="dc2be-135">`Set-App` - Enables or disables a add-in on a mailbox.</span></span>
-  <span data-ttu-id="dc2be-136">`Remove-App`: удаляет ранее установленную надстройку с сервера Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="dc2be-136">`Remove-App` - Removes a previously installed add-in from an Exchange server.</span></span>

## <a name="client-versions"></a><span data-ttu-id="dc2be-137">Версии клиента</span><span class="sxs-lookup"><span data-stu-id="dc2be-137">Client versions</span></span>

<span data-ttu-id="dc2be-138">Выбор версии клиента Outlook для тестирования зависит от ваших требований к разработке.</span><span class="sxs-lookup"><span data-stu-id="dc2be-138">Deciding what versions of the Outlook client to test depends on your development requirements.</span></span>

- <span data-ttu-id="dc2be-p105">Если вы разрабатываете надстройку для частного использования или только для членов организации, важно протестировать версии Outlook, используемые в компании. Обратите внимание, что некоторые пользователи могут использовать Outlook в Интернете, поэтому также важно протестировать версии стандартного браузера компании.</span><span class="sxs-lookup"><span data-stu-id="dc2be-p105">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span></span>

- <span data-ttu-id="dc2be-p106">Если вы разрабатываете надстройку для размещения в [AppSource](https://appsource.microsoft.com), необходимо протестировать версии, указанные в [политиках сертификации коммерческой платформы Marketplace 1120.3](/legal/marketplace/certification-policies#11203-functionality), в том числе:</span><span class="sxs-lookup"><span data-stu-id="dc2be-p106">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:</span></span>
    - <span data-ttu-id="dc2be-143">Последнюю и предпоследнюю версии Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="dc2be-143">The latest version of Outlook on Windows and the version prior to the latest.</span></span>
    - <span data-ttu-id="dc2be-144">Последнюю версию Outlook для Mac.</span><span class="sxs-lookup"><span data-stu-id="dc2be-144">The latest version of Outlook on Mac.</span></span>
    - <span data-ttu-id="dc2be-145">Последнюю версию Outlook для iOS и Android (если надстройка [поддерживает мобильный формат](add-mobile-support.md)).</span><span class="sxs-lookup"><span data-stu-id="dc2be-145">The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).</span></span>
    - <span data-ttu-id="dc2be-146">Версии браузеров, указанные в политике проверки коммерческой платформы Marketplace 1120.3.</span><span class="sxs-lookup"><span data-stu-id="dc2be-146">The browser versions specified in the Commercial marketplace validation policy 1120.3.</span></span>

> [!NOTE]
> <span data-ttu-id="dc2be-147">Если ваша надстройка не поддерживает один из указанных выше клиентов, так как [запрашивает набор обязательных элементов API](apis.md), не поддерживаемый клиентом, его тестировать не нужно.</span><span class="sxs-lookup"><span data-stu-id="dc2be-147">If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.</span></span>

## <a name="see-also"></a><span data-ttu-id="dc2be-148">См. также</span><span class="sxs-lookup"><span data-stu-id="dc2be-148">See also</span></span>

- [<span data-ttu-id="dc2be-149">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="dc2be-149">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
