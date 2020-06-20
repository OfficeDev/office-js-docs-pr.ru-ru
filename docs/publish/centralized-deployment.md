---
title: Публикация надстроек Office с использованием централизованного развертывания в Центре администрирования Office 365
description: Узнайте, как с помощью централизованного развертывания развертывать внутренние надстройки, а также надстройки, предоставляемые поставщиками программного обеспечения.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 1410409fbd86be13da4551b2f140bd41fdaebbbf
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778678"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-office-365-admin-center"></a>Публикация надстроек Office с использованием централизованного развертывания в Центре администрирования Office 365

The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.

В настоящее время центр администрирования Office 365 поддерживает следующие сценарии.

- Централизованное развертывание новых и обновленных надстроек для отдельных пользователей, групп или организации.
- Развертывание на нескольких клиентских платформах, в том числе Windows, Mac и Интернет. Для Outlook также поддерживается развертывание в iOS и Android. (Тем **не** менее, при установке надстроек Excel, Outlook, Word и PowerPoint на iPad не поддерживается централизованное развертывание для iPad.)
- Развертывание на клиентах на английском и других языках.
- Развертывание надстроек, размещаемых в облаке.
- Развертывание надстроек, размещаемых в брандмауэре.
- Развертывание надстроек AppSource.
- Автоматическая установка надстройки для пользователей при запуске приложения Office.
- Автоматическое удаление надстройки для пользователей, если администратор отключит или удалит ее либо пользователь будет удален из службы Azure Active Directory или группы, в которой была развернута надстройка.

Централизованное развертывание — рекомендуемый для администраторов Office 365 способ развертывания надстроек Office в организации при условии, что организация отвечает всем требованиям для использования централизованного развертывания. Сведения о том, как определить, можно ли использовать централизованное развертывание в вашей организации, см. в статье [Оценка соответствия организации Office 365 требованиям для централизованного развертывания надстроек](/office365/admin/manage/centralized-deployment-of-add-ins).

> [!NOTE]
> In an on-premises environment with no connection to Office 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint app catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](/visualstudio/vsto/deploying-an-office-solution).

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Рекомендуемый подход к развертыванию надстроек Office

Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:

1. Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.

2. Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.

3. Разверните надстройку для всех пользователей, которые будут работать с надстройкой.

В зависимости от размера целевой аудитории может потребоваться добавить или убрать этапы этой процедуры.

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>Публикация надстройки Office путем централизованного развертывания

Прежде чем приступать к работе, убедитесь, что организация отвечает всем требованиям для использования централизованного развертывания, как описано в статье [Оценка соответствия организации Office 365 требованиям для централизованного развертывания надстроек](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).

Если организация отвечает всем требованиям, выполните указанные ниже действия, чтобы опубликовать надстройку Office путем централизованного развертывания.

1. Войдите в Office 365 с рабочей или учебной учетной записью.
2. В левом верхнем углу щелкните значок средства запуска приложений и выберите **Администратор**.
3. В меню навигации нажмите **Больше**, а затем выберите **Параметры** > **Службы и надстройки**.
4. Если в верхней части страницы появится сообщение о новом Центре администрирования Office 365, нажмите его, чтобы перейти к ознакомительной версии центра администрирования (см. статью [О Центре администрирования Office 365](/microsoft-365/admin/admin-overview/about-the-admin-center)).
5. В верхней части страницы выберите **Развернуть надстройку**.
6. Просмотрев требования, нажмите кнопку **Далее**.
7. На странице **Централизованное развертывание** выберите один из следующих вариантов:

    - **Я хочу добавить надстройку из Магазина Office**.
    - **I have the manifest file (.xml) on this device.** For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.
    - **I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.

    ![Диалоговое окно создания надстройки в Центре администрирования Office 365](../images/new-add-in.png)

8. Если добавляется надстройка из Магазина Office, выберите ее. Вы можете просматривать доступные надстройки по категориям **Рекомендуемое**, **Оценка** и **Имя**. Из Магазина Office можно добавлять только бесплатные надстройки. Добавление платных надстроек сейчас не поддерживается.

    > [!NOTE]
    > Если выбран вариант с Магазином Office, то обновления и улучшения надстройки автоматически предоставляются пользователям без вашего участия.

    ![Выбор диалогового окна надстройки в центре администрирования Office 365](../images/select-an-add-in.png)

9. Нажмите кнопку **Continue (продолжить** ) после просмотра сведений о надстройках, политики конфиденциальности и условий лицензионного соглашения.

    ![Выбранная страница надстройки в центре администрирования Office 365](../images/selected-add-in-admin-center.png)

10. На странице **Назначение пользователей** выберите **все**, **конкретные пользователи/группы**или **только я**. С помощью поля поиска найдите пользователей и группы, для которых нужно развернуть надстройку. Для надстроек Outlook также можно выбрать метод развертывания **fixed**, **Available**или **Optional**.

    ![Управление пользователями, у которых есть метод доступа и развертывания в центре администрирования Office 365](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > Система [единого входа](../develop/sso-in-office-add-ins.md) для надстроек сейчас доступна в предварительной версии. Ее не следует использовать для рабочих надстроек. При развертывании надстройки с помощью единого входа назначенным пользователям и группам также предоставляется доступ к надстройкам, использующим тот же идентификатор приложения Azure. Все изменения, касающиеся назначений пользователей, также применяются к этим надстройкам. На этой странице отображаются связанные надстройки. На этой странице приводится список разрешений Microsoft Graph, необходимых надстройке (только для надстроек с поддержкой единого входа).

11. По завершении нажмите кнопку **развернуть**. Этот процесс может занять до трех минут. Затем нажмите кнопку **Далее**, чтобы завершить выполнение пошаговых инструкций. Теперь надстройка будет отображаться вместе с другими приложениями в Office 365.

    > [!NOTE]
    > Когда администратор выбирает **развертывание**, согласие предоставляется всем пользователям.

    ![Список приложений в Центре администрирования Office 365](../images/citations.png)

> [!TIP]
> При развертывании новой надстройки для пользователей и/или групп в организации рекомендуем отправлять им электронные сообщения с указаниями по использованию надстройки и ссылками на соответствующие разделы справки, часто задаваемые вопросы и другие вспомогательные ресурсы.

## <a name="considerations-when-granting-access-to-an-add-in"></a>Рекомендации по предоставлению доступа к надстройке

Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:

- **Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.

- **Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.

- **Groups**: If you assign an add-in to a group, users who are added to the group will automatically be assigned the add-in. Likewise, when a user is removed from a group, the user automatically loses access to the add-in. In either case, no additional action is required from the Office 365 admin.

In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.

## <a name="add-in-states"></a>Состояния надстроек

В приведенной ниже таблице описываются различные состояния надстройки.

|Состояние|Причины|Влияние|
|-----|--------------------|------|
|**Активна**|Администратор отправил надстройку и назначил ее пользователям и/или группам.|Надстройка видна назначенным пользователям и/или группам в соответствующих клиентах Office.|
|**Отключена**|Администратор отключил надстройку.|Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.|
|**Deleted**|Администратор удалил надстройку.|Надстройка недоступна назначенным пользователям и/или группам.|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>Обновление надстроек Office, опубликованных с использованием централизованного развертывания

After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users as soon as those changes are implemented in the web application. Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md), for example, to update the add-in's icon, text, or add-in commands, happen as follows:

- **Line-of-business add-in**: If an admin explicitly uploaded a manifest file when implementing Centralized Deployment via the Office 365 admin center, the admin must upload a new manifest file that contains the desired changes. After the updated manifest file has been uploaded, the next time the relevant Office applications start, the add-in will update.

  > [!NOTE]
  > Администратору не нужно удалять бизнес-надстройку, чтобы выполнить обновление. В разделе надстройки администратор может просто выбрать бизнес-надстройку и вызвать эту функцию, нажав кнопку **обновить надстройку** в правом нижнем углу.
  > 
  > ![На снимке экрана отображается диалоговое окно обновления надстройки в центре администрирования Office 365](../images/update-add-in-admin-center.png)

- **Надстройка из Магазина Office**. Если администратор выбрал надстройку из Магазина Office во время реализации централизованного развертывания в Центре администрирования Office 365, а надстройка в Магазине Office обновилась, то она будет обновлена позже с использованием централизованного развертывания. Надстройка обновится при следующем запуске соответствующего приложения Office.

## <a name="end-user-experience-with-add-ins"></a>Работа пользователей с надстройками

После публикации надстройки с применением централизованного развертывания пользователи могут приступить к работе с ней на любой платформе, которую поддерживает надстройка.

If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.

![Снимок экрана с разделом ленты Office, где выделена команда Search Citation (Поиск ссылки) для надстройки Citations (Ссылки)](../images/search-citation.png)

Если команды надстройки не поддерживаются, пользователи могут добавить надстройку в свое приложение Office, сделав вот что:

1. В Word 2016, Excel 2016 или PowerPoint 2016 либо более поздних версий выберите **Вставка** > **Мои надстройки**.
2. В окне надстройки перейдите на вкладку **Управляемые администратором**.
3. Выберите нужную надстройку и нажмите **Добавить**.

    ![Screenshot shows the Admin Managed tab of the Office Add-ins page of an Office application. The Citations add-in is shown on the tab.](../images/office-add-ins-admin-managed.png)

Однако в Outlook 2016 или более поздних версий можно сделать вот что:

1. В Outlook выберите **Главная** > **Магазин**.
2. На вкладке надстройки выберите элемент **Управляемые администратором**.
3. Выберите нужную надстройку и нажмите **Добавить**.

    ![Снимок экрана с областью "Управляемые администратором" страницы "Магазин" в приложении Outlook.](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>См. также

- [Определение пригодности централизованного развертывания надстроек для вашей организации Office 365](/office365/admin/manage/centralized-deployment-of-add-ins)
