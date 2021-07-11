---
title: Тестирование удобства использования надстроек Office
description: Узнайте, как проверить дизайн надстройки с реальными пользователями.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 5de29a15a9e382b990985765eaad801b1b54f364
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349695"
---
# <a name="usability-testing-for-office-add-ins"></a><span data-ttu-id="31aa7-103">Тестирование удобства использования надстроек Office</span><span class="sxs-lookup"><span data-stu-id="31aa7-103">Usability testing for Office Add-ins</span></span>

<span data-ttu-id="31aa7-p101">Для создания качественной надстройки необходимо учитывать поведение пользователей. Так как предубеждения разработчиков влияют на принятие проектных решений, важно тестировать надстройки с настоящими пользователями, чтобы гарантировать хорошую работу надстройки в реальных ситуациях.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p101">A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.</span></span> 

<span data-ttu-id="31aa7-106">Вы можете запускать тесты на доступность по-разному.</span><span class="sxs-lookup"><span data-stu-id="31aa7-106">You can run usability tests in different ways.</span></span> <span data-ttu-id="31aa7-107">Для многих разработчиков надстройки исследования удаленного и неизмеримого использования являются наиболее эффективными по времени и затратам.</span><span class="sxs-lookup"><span data-stu-id="31aa7-107">For many add-in developers, remote, unmoderated usability studies are the most time and cost effective.</span></span> <span data-ttu-id="31aa7-108">Это легко сделать несколькими популярными службами тестирования. Ниже приводится несколько примеров.</span><span class="sxs-lookup"><span data-stu-id="31aa7-108">Several popular testing services make this easy; the following are some examples.</span></span>

- [<span data-ttu-id="31aa7-109">UserTesting.com</span><span class="sxs-lookup"><span data-stu-id="31aa7-109">UserTesting.com</span></span>](https://www.UserTesting.com)
- [<span data-ttu-id="31aa7-110">Optimalworkshop.com</span><span class="sxs-lookup"><span data-stu-id="31aa7-110">Optimalworkshop.com</span></span>](https://www.Optimalworkshop.com)
- [<span data-ttu-id="31aa7-111">Userzoom.com</span><span class="sxs-lookup"><span data-stu-id="31aa7-111">Userzoom.com</span></span>](https://www.Userzoom.com)

<span data-ttu-id="31aa7-112">Эти службы тестирования помогают упростить создание плана тестирования, а также избавляют от необходимости искать участников и наблюдать за тестированием.</span><span class="sxs-lookup"><span data-stu-id="31aa7-112">These testing services help you to streamline test plan creation and remove the need to seek out participants or moderate the tests.</span></span> 

<span data-ttu-id="31aa7-p103">Пяти участников достаточно, чтобы обнаружить большую часть проблем при использовании надстройки. В течение цикла разработки регулярно проводите небольшие испытания, чтобы убедиться, что в вашем продукте учитываются потребности пользователей.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p103">You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.</span></span>

> [!NOTE]
> <span data-ttu-id="31aa7-p104">Рекомендуем тестировать удобство использования надстроек на нескольких платформах. Для [публикации надстройки в AppSource](/office/dev/store/submit-to-appsource-via-partner-center) она должна работать на всех [платформах, поддерживающих определенные вами методы](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="31aa7-p104">We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center), it must work on all [platforms that support the methods that you define](../overview/office-add-in-availability.md).</span></span>

## <a name="1---sign-up-for-a-testing-service"></a><span data-ttu-id="31aa7-117">1. Зарегистрируйтесь в службе тестирования</span><span class="sxs-lookup"><span data-stu-id="31aa7-117">1.   Sign up for a testing service</span></span>

<span data-ttu-id="31aa7-118">Дополнительные сведения см. в статье [Выбор веб-инструмента для удаленного немодерируемого тестирования](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).</span><span class="sxs-lookup"><span data-stu-id="31aa7-118">For more information, see [Selecting an Online Tool for Unmoderated Remote User Testing](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).</span></span>

## <a name="2-develop-your-research-questions"></a><span data-ttu-id="31aa7-119">2. Определите предметы исследования</span><span class="sxs-lookup"><span data-stu-id="31aa7-119">2. Develop your research questions</span></span>
 
<span data-ttu-id="31aa7-p105">Предметы исследования определяют цели и план тестирования. Они помогут вам выбрать участников и назначить им задачи. Предметы исследования должны быть как можно более конкретными. Вы также можете поставить общие вопросы.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p105">Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.</span></span>
 
<span data-ttu-id="31aa7-124">Ниже приводится несколько примеров исследовательских вопросов.</span><span class="sxs-lookup"><span data-stu-id="31aa7-124">The following are some examples of research questions.</span></span>
  
<span data-ttu-id="31aa7-125">**Конкретные**</span><span class="sxs-lookup"><span data-stu-id="31aa7-125">**Specific**</span></span>

- <span data-ttu-id="31aa7-126">Замечают ли пользователи ссылку "Бесплатная пробная версия" на целевой странице?</span><span class="sxs-lookup"><span data-stu-id="31aa7-126">Do users notice the "free trial" link on the landing page?</span></span>
- <span data-ttu-id="31aa7-127">Когда пользователи вставляют содержимое из надстройки в документ, знают ли они, в каком месте документа оно будет вставлено?</span><span class="sxs-lookup"><span data-stu-id="31aa7-127">When users insert content from the add-in to their document, do they understand where in the document it is inserted?</span></span>

<span data-ttu-id="31aa7-128">**Общие**</span><span class="sxs-lookup"><span data-stu-id="31aa7-128">**Broad**</span></span>

- <span data-ttu-id="31aa7-129">С чем у пользователя возникает больше всего сложностей при работе с надстройкой?</span><span class="sxs-lookup"><span data-stu-id="31aa7-129">What are the biggest pain points for the user in our add-in?</span></span>
- <span data-ttu-id="31aa7-130">Понимают ли пользователи значения значков на панели команд, прежде чем нажимать их?</span><span class="sxs-lookup"><span data-stu-id="31aa7-130">Do users understand the meaning of the icons in our command bar, before they click on them?</span></span>
- <span data-ttu-id="31aa7-131">Легко ли пользователям найти меню настроек?</span><span class="sxs-lookup"><span data-stu-id="31aa7-131">Can users easily find the settings menu?</span></span>

<span data-ttu-id="31aa7-132">Очень важно собрать данные обо всем процессе работы пользователя — от обнаружения надстройки до ее установки и использования.</span><span class="sxs-lookup"><span data-stu-id="31aa7-132">It’s important to get data on the entire user journey – from discovering your add-in, to installing and using it.</span></span> <span data-ttu-id="31aa7-133">Рассмотрите вопросы исследований, которые рассматривают следующие аспекты пользовательского интерфейса надстройки.</span><span class="sxs-lookup"><span data-stu-id="31aa7-133">Consider research questions that address the following aspects of the add-in user experience.</span></span>

- <span data-ttu-id="31aa7-134">поиск надстройки в AppSource;</span><span class="sxs-lookup"><span data-stu-id="31aa7-134">Finding your add-in in AppSource</span></span>
- <span data-ttu-id="31aa7-135">установка надстройки;</span><span class="sxs-lookup"><span data-stu-id="31aa7-135">Choosing to install your add-in</span></span>
- <span data-ttu-id="31aa7-136">первый запуск;</span><span class="sxs-lookup"><span data-stu-id="31aa7-136">First run experience</span></span>
- <span data-ttu-id="31aa7-137">команды ленты;</span><span class="sxs-lookup"><span data-stu-id="31aa7-137">Ribbon commands</span></span>
- <span data-ttu-id="31aa7-138">интерфейс надстройки;</span><span class="sxs-lookup"><span data-stu-id="31aa7-138">Add-in UI</span></span>
- <span data-ttu-id="31aa7-139">взаимодействие надстройки с пространством документа в приложении Office;</span><span class="sxs-lookup"><span data-stu-id="31aa7-139">How the add-in interacts with the document space of the Office application</span></span>
- <span data-ttu-id="31aa7-140">возможности пользователя по управлению вставкой содержимого.</span><span class="sxs-lookup"><span data-stu-id="31aa7-140">How much control the user has over any content insertion flows</span></span>

<span data-ttu-id="31aa7-141">Дополнительные сведения см. в статье [Gathering factual responses vs. subjective data](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions) (Сбор фактических ответов и субъективных данных).</span><span class="sxs-lookup"><span data-stu-id="31aa7-141">For more information, see [Gathering factual responses vs. subjective data](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions).</span></span>

## <a name="3-identify-participants-to-target"></a><span data-ttu-id="31aa7-142">3. Определите участников тестирования</span><span class="sxs-lookup"><span data-stu-id="31aa7-142">3. Identify participants to target</span></span>

<span data-ttu-id="31aa7-p107">Удаленные службы тестирования позволяют контролировать множество характеристик участников тестирования. Задумайтесь, для каких пользователей предназначено исследование. На ранних этапах сбора данных рекомендуем набрать широкую аудиторию участников, чтобы выявить наиболее очевидные проблемы при использовании. Позже вам могут потребоваться конкретные целевые группы, например опытные пользователи Office, определенные профессии или возрастные категории.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p107">Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.</span></span>

## <a name="4-create-the-participant-screener"></a><span data-ttu-id="31aa7-147">4. Создайте анкету для отбора участников</span><span class="sxs-lookup"><span data-stu-id="31aa7-147">4. Create the participant screener</span></span>

<span data-ttu-id="31aa7-p108">Критерии отбора — это набор вопросов и требований, которые вы представляете потенциальным участникам тестирования. Помните, что участники в таких службах, как UserTesting.com, материально заинтересованы участвовать в вашем тестировании. Если вы хотите исключить участие определенных пользователей, рекомендуем задать несколько каверзных вопросов.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p108">The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test.</span></span> 

<span data-ttu-id="31aa7-151">Например, если вы хотите найти участников, знакомых с сайтом GitHub, добавьте в список вариантов ответа вымышленные названия, чтобы отфильтровать пользователей, предоставляющих ложные сведения.</span><span class="sxs-lookup"><span data-stu-id="31aa7-151">For example, if you want to find participants who are familiar with GitHub, to filter out users who might misrepresent themselves, include fakes in the list of possible answers.</span></span>

<span data-ttu-id="31aa7-152">**С каким из перечисленных ниже репозиториев кода вы знакомы?**</span><span class="sxs-lookup"><span data-stu-id="31aa7-152">**Which of the following source code repositories are you familiar with?**</span></span>  
 <span data-ttu-id="31aa7-p109">А. SourceShelf  [*Отклонить*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p109">a. SourceShelf  [*Reject*]</span></span>  
 <span data-ttu-id="31aa7-p110">Б. CodeContainer  [*Отклонить*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p110">b. CodeContainer  [*Reject*]</span></span>  
 <span data-ttu-id="31aa7-p111">В. GitHub  [*Необходимо выбрать*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p111">c. GitHub  [*Must select*]</span></span>  
 <span data-ttu-id="31aa7-p112">Г. BitBucket  [*Можно выбрать*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p112">d. BitBucket  [*May select*]</span></span>  
 <span data-ttu-id="31aa7-p113">Д. CloudForge  [*Можно выбрать*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p113">e. CloudForge  [*May select*]</span></span>  

<span data-ttu-id="31aa7-163">Если вы планируете тестировать готовую сборку надстройки, представленные ниже вопросы помогут вам отобрать пользователей, которые смогут это сделать.</span><span class="sxs-lookup"><span data-stu-id="31aa7-163">If you are planning to test a live build of your add-in, the following questions can screen for users who will be able to do this.</span></span>

<span data-ttu-id="31aa7-164">**Для этого теста вам потребуется последняя версия Microsoft PowerPoint. У вас есть последняя версия PowerPoint?**</span><span class="sxs-lookup"><span data-stu-id="31aa7-164">**This test requires you to have the latest version of Microsoft PowerPoint. Do you have the latest version of PowerPoint?**</span></span>  
 <span data-ttu-id="31aa7-p114">a. Да [*Должен выбрать*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p114">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="31aa7-p115">b. Нет [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p115">b. No [*Reject*]</span></span>  
 <span data-ttu-id="31aa7-p116">c. Не знаю [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p116">c. I don’t know [*Reject*]</span></span>  

<span data-ttu-id="31aa7-171">**Для этого теста вам потребуется установить бесплатную надстройку для PowerPoint и создать бесплатную учетную запись для ее использования. Вы готовы установить надстройку и создать бесплатную учетную запись?**</span><span class="sxs-lookup"><span data-stu-id="31aa7-171">**This test requires you to install a free add-in for PowerPoint, and create a free account to use it. Are you willing to install an add-in and create a free account?**</span></span>  
 <span data-ttu-id="31aa7-p117">a. Да [*Должен выбрать*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p117">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="31aa7-p118">b. Нет [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="31aa7-p118">b. No [*Reject*]</span></span>  

<span data-ttu-id="31aa7-176">Дополнительные сведения см. в статье [Рекомендации по составлению вопросов для отбора](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices).</span><span class="sxs-lookup"><span data-stu-id="31aa7-176">For more information, see [Screener Questions Best Practices](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices).</span></span>

## <a name="5-create-tasks-and-questions-for-participants"></a><span data-ttu-id="31aa7-177">5. Составьте список задач и вопросов для участников</span><span class="sxs-lookup"><span data-stu-id="31aa7-177">5. Create tasks and questions for participants</span></span>

<span data-ttu-id="31aa7-p119">Постарайтесь расставить приоритеты тестирования, чтобы ограничить количество задач и вопросов для каждого участника. Некоторые службы платят участникам только за определенное время, поэтому выполнение ваших заданий не должно занимать слишком много времени.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p119">Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.</span></span>

<span data-ttu-id="31aa7-p120">По возможности старайтесь наблюдать за действиями участников, а не спрашивать о них. Если вам нужно задавать вопросы о работе пользователей, спросите, что участники делали в таких ситуациях раньше, а не об их предположениях. Как правило, это дает более надежные результаты.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p120">Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.</span></span>

<span data-ttu-id="31aa7-p121">Самая сложная задача при тестировании без наблюдения — убедиться, что участники понимают ваши задания и контекст работы. Указания должны быть *понятными и краткими*. В неоднозначных указаниях кто-то обязательно запутается.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p121">The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.</span></span>

<span data-ttu-id="31aa7-p122">Не полагайтесь на то, что у пользователя будет открыт именно тот экран, который необходим на этом этапе тестирования. Рекомендуем сообщать участникам, какой экран должен быть открыт в начале следующей задачи.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p122">Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.</span></span>

<span data-ttu-id="31aa7-188">Дополнительные сведения см. в статье [Составление хороших задач](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks).</span><span class="sxs-lookup"><span data-stu-id="31aa7-188">For more information, see [Writing Great Tasks](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks).</span></span>

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a><span data-ttu-id="31aa7-189">6. Создайте прототип, соответствующий задачам и вопросам</span><span class="sxs-lookup"><span data-stu-id="31aa7-189">6. Create a prototype to match the tasks and questions</span></span>
 
<span data-ttu-id="31aa7-190">Вы можете тестировать либо готовую версию надстройки, либо ее прототип.</span><span class="sxs-lookup"><span data-stu-id="31aa7-190">You can either test your live add-in, or you can test a prototype.</span></span> <span data-ttu-id="31aa7-191">Помните, что если вы хотите протестировать готовую версию надстройки, вам потребуется отобрать участников, у которых есть последняя версия Office и которые готовы установить надстройку и зарегистрировать учетную запись (если у вас нет для них готовых учетных данных для входа). Затем необходимо убедиться, что они успешно установили надстройку.</span><span class="sxs-lookup"><span data-stu-id="31aa7-191">Keep in mind that if you want to test the live add-in, you need to screen for participants that have the latest version of Office, are willing to install the add-in, and are willing to sign up for an account (unless you have logon credentials to provide them.) You'll then need to make sure that they successfully install your add-in.</span></span>

<span data-ttu-id="31aa7-p124">В среднем обучение пользователей установке надстройки занимает около 5 минут. Ниже представлен пример понятных, кратких указаний по установке. Откорректируйте их в соответствии с особенностями вашего тестирования.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p124">On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.</span></span>

<span data-ttu-id="31aa7-195">**Установите надстройки (вставьте имя надстройки здесь) для PowerPoint, используя следующие инструкции.**</span><span class="sxs-lookup"><span data-stu-id="31aa7-195">**Please install the (insert your add-in name here) add-in for PowerPoint, using the following instructions.**</span></span>

1. <span data-ttu-id="31aa7-196">Откройте Microsoft PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="31aa7-196">Open Microsoft PowerPoint.</span></span>
1. <span data-ttu-id="31aa7-197">Выберите элемент **Пустая презентация**.</span><span class="sxs-lookup"><span data-stu-id="31aa7-197">Select **Blank Presentation.**</span></span>
1. <span data-ttu-id="31aa7-198">Нажмите **Вставка > Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="31aa7-198">Go to **Insert > My Add-ins.**</span></span>
1. <span data-ttu-id="31aa7-199">Во всплывающем окне выберите **Магазин**.</span><span class="sxs-lookup"><span data-stu-id="31aa7-199">In the popup window, choose **Store.**</span></span>
1. <span data-ttu-id="31aa7-200">Введите <название надстройки> в поле поиска.</span><span class="sxs-lookup"><span data-stu-id="31aa7-200">Type (Add-in name) in the search box.</span></span>
1. <span data-ttu-id="31aa7-201">Выберите <название надстройки>.</span><span class="sxs-lookup"><span data-stu-id="31aa7-201">Choose (Add-in name).</span></span>
1. <span data-ttu-id="31aa7-202">Изучите страницу Магазина, чтобы ознакомиться с надстройкой.</span><span class="sxs-lookup"><span data-stu-id="31aa7-202">Take a moment to look at the Store page to familiarize yourself with the add-in.</span></span>
1. <span data-ttu-id="31aa7-203">Нажмите **Добавить**, чтобы установить надстройку.</span><span class="sxs-lookup"><span data-stu-id="31aa7-203">Choose **Add** to install the add-in.</span></span>

<span data-ttu-id="31aa7-p125">Вы можете тестировать прототип на любом уровне взаимодействия и визуальной четкости. Для создания более сложных связей и повышенной интерактивности можно использовать такие средства создания прототипов, как [InVision](https://www.invisionapp.com). Если требуется протестировать статические экраны, вы можете разместить изображения в Интернете и отправить участникам соответствующий URL-адрес либо предоставить им ссылку на презентацию PowerPoint в Интернете.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p125">You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation.</span></span> 

## <a name="7-run-a-pilot-test"></a><span data-ttu-id="31aa7-207">7. Проведите пилотный тест</span><span class="sxs-lookup"><span data-stu-id="31aa7-207">7. Run a pilot test</span></span>

<span data-ttu-id="31aa7-p126">Создание прототипа и составление списка задач или вопросов могут быть непростыми задачами. Пользователи могут запутаться в заданиях или не разобраться в прототипе. Следует провести пилотный тест с 1–3 пользователями, чтобы выявить неизбежные проблемы с форматом тестирования. Это поможет гарантировать, что ваши вопросы ясны, прототип настроен правильно и вы собираете именно те данные, которые вам нужны.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p126">It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.</span></span>

## <a name="8-run-the-test"></a><span data-ttu-id="31aa7-212">8. Проведите тестирование</span><span class="sxs-lookup"><span data-stu-id="31aa7-212">8. Run the test</span></span>

<span data-ttu-id="31aa7-p127">Заказав тест, вы будете получать по электронной почте уведомления, когда участники проходят его. Как правило, если у вас нет специфических требований к участникам, тесты завершаются в течение нескольких часов.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p127">After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.</span></span>

## <a name="9-analyze-results"></a><span data-ttu-id="31aa7-215">9. Проанализируйте результаты</span><span class="sxs-lookup"><span data-stu-id="31aa7-215">9. Analyze results</span></span>

<span data-ttu-id="31aa7-p128">На этом этапе вам необходимо сделать выводы из собранных данных. Просматривая видео тестирования, делайте заметки о проблемах и успехах пользователя. Старайтесь не делать поспешных выводов. Просмотрите все результаты.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p128">This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.</span></span> 

<span data-ttu-id="31aa7-p129">Проблема, возникшая у одного участника, еще не означает, что необходимо что-то менять. Если же одна и та же проблема возникнет у нескольких участников, значит, она возникнет и у других пользователей в широкой аудитории.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p129">A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.</span></span>

<span data-ttu-id="31aa7-p130">В целом, будьте осторожны с выводами при анализе данных. Не совершайте ошибку, стараясь подогнать результаты под определенную интерпретацию. Объективно оцените, что полученные данные доказывают или опровергают, а для каких выводов их просто недостаточно. Избегайте предвзятости — поведение пользователей часто противоречит ожиданиям разработчика.</span><span class="sxs-lookup"><span data-stu-id="31aa7-p130">In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.</span></span>


## <a name="see-also"></a><span data-ttu-id="31aa7-224">См. также</span><span class="sxs-lookup"><span data-stu-id="31aa7-224">See also</span></span>

- [<span data-ttu-id="31aa7-225">Как проводить тестирование удобства использования</span><span class="sxs-lookup"><span data-stu-id="31aa7-225">How to Conduct Usability Testing</span></span>](https://whatpixel.com/howto-conduct-usability-testing/)
- [<span data-ttu-id="31aa7-226">Рекомендации по пользовательскому тестированию</span><span class="sxs-lookup"><span data-stu-id="31aa7-226">Best Practices for UserTesting</span></span>](https://help.usertesting.com/hc/articles/115003370231-Best-practices-for-UserTesting)  
- [<span data-ttu-id="31aa7-227">Предотвращение предвзятости</span><span class="sxs-lookup"><span data-stu-id="31aa7-227">Minimizing Bias</span></span>](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)
