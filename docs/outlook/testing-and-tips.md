---
title: Развертывание и установка надстроек Outlook для тестирования
description: Создайте файл манифеста, разверните файл пользовательского интерфейса надстройки на веб-сервере, установите надстройку в своем почтовом ящике, а затем протестируйте ее.
ms.date: 05/20/2020
localization_priority: Priority
ms.openlocfilehash: 97841f7c8112b42cee2927f238b31fe985b2e101
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093863"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>Развертывание и установка надстроек Outlook для тестирования

В рамках разработки надстройки Outlook вам, скорее всего, понадобится несколько раз развертывать и устанавливать надстройку для тестирования, что подразумевает выполнение следующих действий.

1. Создание файла манифеста, в котором описывается надстройка.
1. Развертывание файлов пользовательского интерфейса надстройки на веб-сервере.
1. Установка надстройки в почтовом ящике пользователя.
1. Тестирование надстройки с внесением соответствующих изменений в пользовательский интерфейс или файлы манифеста и повторение этапов 2 и 3 для тестирования изменений.

> [!NOTE]
> Поскольку [настраиваемые области устарели](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), следует убедиться, что вы используете [поддерживаемую точку расширения надстройки](outlook-add-ins-overview.md#extension-points).

## <a name="create-a-manifest-file-for-the-add-in"></a>Создание файла манифеста для надстройки

Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).

## <a name="deploy-an-add-in-to-a-web-server"></a>Развертывание надстройки на веб-сервере

You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.

## <a name="install-the-add-in"></a>Установка надстройки

После подготовки файла манифеста и развертывания пользовательского интерфейса надстройки на доступном веб-сервере, вы можете загрузить неопубликованную надстройку для почтового ящика на сервере Exchange Server, используя клиент Outlook, или установить ее с помощью командлетов Windows PowerShell.

### <a name="sideload-the-add-in"></a>Загрузка неопубликованной надстройки

You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.

The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

При необходимости администратор может выполнить следующий командлет, чтобы назначить похожие разрешения нескольким пользователям:

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

Дополнительные сведения об упомянутой роли см. в статье [Роль My Custom Apps](/exchange/my-custom-apps-role-exchange-2013-help).

Если для разработки надстроек вы используете Microsoft 365 или Visual Studio, вам назначается роль администратора организации, позволяющая устанавливать надстройки с помощью файла или URL-адреса в Центре администрирования Exchange, а также с помощью командлетов PowerShell.

### <a name="install-an-add-in-by-using-remote-powershell"></a>Установка надстройки с помощью удаленного сеанса PowerShell

После создания удаленного сеанса Windows PowerShell на сервере Exchange Server вы можете установить надстройку Outlook, используя командлет `New-App` и следующую команду PowerShell.

```powershell
New-App -URL:"http://<fully-qualified URL">
```

Полный URL-адрес — это расположение подготовленного файла манифеста надстройки.

Вы можете использовать следующие командлеты PowerShell для управления надстройками для почтового ящика:

- `Get-App`: отображает надстройки, включенные для почтового ящика.
- `Set-App`: включает или отключает надстройку для почтового ящика.
- `Remove-App`: удаляет ранее установленную надстройку с сервера Exchange Server.

## <a name="client-versions"></a>Версии клиента

Выбор версии клиента Outlook для тестирования зависит от ваших требований к разработке.

- If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.

- If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:
  - Последнюю и предпоследнюю версии Outlook для Windows.
  - Последнюю версию Outlook для Mac.
  - Последнюю версию Outlook для iOS и Android (если надстройка [поддерживает мобильный формат](add-mobile-support.md)).
  - Версии браузеров, указанные в политике проверки коммерческой платформы Marketplace 1120.3.

> [!NOTE]
> Если ваша надстройка не поддерживает один из указанных выше клиентов, так как [запрашивает набор обязательных элементов API](apis.md), не поддерживаемый клиентом, его тестировать не нужно.

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Outlook в Интернете и версии Exchange Server

Потребители и пользователи учетной записи Microsoft 365 видят современную версию интерфейса при обращении к Outlook в Интернете и больше не видят классическую версию, поддержка которой прекращена. Однако локальные серверы Exchange Server продолжают поддерживать классическую версию Outlook в Интернете. Поэтому во время проверки ваша отправка может получить предупреждение о том, что надстройка несовместима с классической версией Outlook в Интернете. В этом случае рекомендуется проверить надстройку в локальной среде Exchange. При этом предупреждение не блокирует отправку в AppSource, но для ваших пользователей могут быть ограничены возможности, если они используют Outlook в Интернете в локальной среде Exchange.

Чтобы устранить эту проблему, рекомендуем проверить надстройку в Outlook в Интернете, подключенном к собственной приватной локальной среде Exchange. Дополнительные сведения см. в руководстве о том, как [создать тестовую среду Exchange 2016 или Exchange 2019](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019#establish-an-exchange-2016-or-exchange-2019-test-environment), и о том, как управлять [Outlook в Интернете в Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019).

Вы также можете выбрать вариант с оплатой и использовать службу, размещающую локальные серверы Exchange Server и управляющую ими. Несколько вариантов:

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/products-services-2/hosted-microsoft-exchange/)

Кроме того, если вы не хотите, чтобы ваши надстройки были доступны для пользователей, подключенных к локальной среде Exchange, вы можете настроить для [набора обязательных элементов](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support) в манифесте надстройки версию 1.6 или более позднюю. Такие надстройки не будут проверяться в классическом интерфейсе Outlook в Интернете.

## <a name="see-also"></a>См. также

- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
