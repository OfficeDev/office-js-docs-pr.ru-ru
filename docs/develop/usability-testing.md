---
title: Тестирование удобства использования надстроек Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 410b8d7ede22cf222ee2df794e438c7f5f8881dd
ms.sourcegitcommit: 4e4f7c095e8f33b06bd8a02534ee901125eb1d17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/28/2018
ms.locfileid: "20085488"
---
# <a name="usability-testing-for-office-add-ins"></a><span data-ttu-id="c84a1-102">Тестирование удобства использования надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c84a1-102">Usability testing for Office Add-ins</span></span>

<span data-ttu-id="c84a1-p101">Для создания качественной надстройки необходимо учитывать поведение пользователей. Так как предубеждения разработчиков влияют на принятие проектных решений, важно тестировать надстройки с настоящими пользователями, чтобы гарантировать хорошую работу надстройки в реальных ситуациях.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p101">A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.</span></span> 

<span data-ttu-id="c84a1-p102">Тестировать удобство использования можно различными способами. Для многих разработчиков надстроек наиболее эффективными и экономичными являются испытания без наблюдения. Существует ряд популярных служб, значительно упрощающих проведение таких испытаний. К ним относятся:</span><span class="sxs-lookup"><span data-stu-id="c84a1-p102">You can run usability tests in different ways. For many add-in developers, remote, unmoderated usability studies are the most time and cost effective. Several popular testing services make this easy; the following are some examples:</span></span> 

 - [<span data-ttu-id="c84a1-108">UserTesting.com</span><span class="sxs-lookup"><span data-stu-id="c84a1-108">UserTesting.com</span></span>](https://www.UserTesting.com)
 - [<span data-ttu-id="c84a1-109">Optimalworkshop.com</span><span class="sxs-lookup"><span data-stu-id="c84a1-109">Optimalworkshop.com</span></span>](https://www.Optimalworkshop.com)
 - [<span data-ttu-id="c84a1-110">Userzoom.com</span><span class="sxs-lookup"><span data-stu-id="c84a1-110">Userzoom.com</span></span>](https://www.Userzoom.com)

<span data-ttu-id="c84a1-111">Эти службы тестирования помогают упростить создание плана тестирования, а также избавляют от необходимости искать участников и наблюдать за тестированием.</span><span class="sxs-lookup"><span data-stu-id="c84a1-111">These testing services help you to streamline test plan creation and remove the need to seek out participants or moderate the tests.</span></span> 

<span data-ttu-id="c84a1-p103">Пяти участников достаточно, чтобы обнаружить большую часть проблем при использовании надстройки. В течение цикла разработки регулярно проводите небольшие испытания, чтобы убедиться, что в вашем продукте учитываются потребности пользователей.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p103">You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.</span></span>

> [!NOTE]
> <span data-ttu-id="c84a1-p104">Рекомендуем тестировать удобство использования надстроек на нескольких платформах. Для [публикации надстройки в AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store) она должна работать на всех [платформах, поддерживающих определенные вами методы](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="c84a1-p104">We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store), it must work on all [platforms that support the methods that you define](../overview/office-add-in-availability.md).</span></span>

## <a name="1---sign-up-for-a-testing-service"></a><span data-ttu-id="c84a1-116">1. Зарегистрируйтесь в службе тестирования</span><span class="sxs-lookup"><span data-stu-id="c84a1-116">1.   Sign up for a testing service</span></span>

<span data-ttu-id="c84a1-117">Дополнительные сведения см. в статье [Выбор веб-инструмента для удаленного немодерируемого тестирования.](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)</span><span class="sxs-lookup"><span data-stu-id="c84a1-117">For more information, see [Selecting an Online Tool for Unmoderated Remote User Testing.](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)</span></span>

## <a name="2-develop-your-research-questions"></a><span data-ttu-id="c84a1-118">2. Определите предметы исследования</span><span class="sxs-lookup"><span data-stu-id="c84a1-118">2. Develop your research questions</span></span>
 
<span data-ttu-id="c84a1-p105">Предметы исследования определяют цели и план тестирования. Они помогут вам выбрать участников и назначить им задачи. Предметы исследования должны быть как можно более конкретными. Вы также можете поставить общие вопросы.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p105">Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.</span></span>
 
<span data-ttu-id="c84a1-123">Ниже приводятся примеры предметов исследования.</span><span class="sxs-lookup"><span data-stu-id="c84a1-123">The following are some examples of research questions:</span></span>
  
<span data-ttu-id="c84a1-124">**Конкретные**</span><span class="sxs-lookup"><span data-stu-id="c84a1-124">**Specific**</span></span>  

 - <span data-ttu-id="c84a1-125">Замечают ли пользователи ссылку "Бесплатная пробная версия" на целевой странице?</span><span class="sxs-lookup"><span data-stu-id="c84a1-125">Do users notice the "free trial" link on the landing page?</span></span>
 - <span data-ttu-id="c84a1-126">Когда пользователи вставляют содержимое из надстройки в документ, знают ли они, в каком месте документа оно будет вставлено?</span><span class="sxs-lookup"><span data-stu-id="c84a1-126">When users insert content from the add-in to their document, do they understand where in the document it is inserted?</span></span>

<span data-ttu-id="c84a1-127">**Общие**</span><span class="sxs-lookup"><span data-stu-id="c84a1-127">**Broad**</span></span>  

 - <span data-ttu-id="c84a1-128">С чем у пользователя возникает больше всего сложностей при работе с надстройкой?</span><span class="sxs-lookup"><span data-stu-id="c84a1-128">What are the biggest pain points for the user in our add-in?</span></span>
 - <span data-ttu-id="c84a1-129">Понимают ли пользователи значения значков на панели команд, прежде чем нажимать их?</span><span class="sxs-lookup"><span data-stu-id="c84a1-129">Do users understand the meaning of the icons in our command bar, before they click on them?</span></span>
 - <span data-ttu-id="c84a1-130">Легко ли пользователям найти меню настроек?</span><span class="sxs-lookup"><span data-stu-id="c84a1-130">Can users easily find the settings menu?</span></span>

<span data-ttu-id="c84a1-p106">Очень важно собрать данные обо всем процессе работы пользователя — от обнаружения надстройки до ее установки и использования. Выберите предметы исследования, относящиеся к следующим аспектам взаимодействия с пользователем:</span><span class="sxs-lookup"><span data-stu-id="c84a1-p106">It’s important to get data on the entire user journey – from discovering your add-in, to installing and using it. Consider research questions that address the following aspects of the add-in user experience:</span></span>
 
 - <span data-ttu-id="c84a1-133">поиск надстройки в AppSource;</span><span class="sxs-lookup"><span data-stu-id="c84a1-133">Finding your add-in in AppSource</span></span>
 - <span data-ttu-id="c84a1-134">установка надстройки;</span><span class="sxs-lookup"><span data-stu-id="c84a1-134">Choosing to install your add-in</span></span>
 - <span data-ttu-id="c84a1-135">первый запуск;</span><span class="sxs-lookup"><span data-stu-id="c84a1-135">First run experience</span></span>
 - <span data-ttu-id="c84a1-136">команды ленты;</span><span class="sxs-lookup"><span data-stu-id="c84a1-136">Ribbon commands</span></span>
 - <span data-ttu-id="c84a1-137">интерфейс надстройки;</span><span class="sxs-lookup"><span data-stu-id="c84a1-137">Add-in UI</span></span>
 - <span data-ttu-id="c84a1-138">взаимодействие надстройки с пространством документа в приложении Office;</span><span class="sxs-lookup"><span data-stu-id="c84a1-138">How the add-in interacts with the document space of the Office application</span></span>
 - <span data-ttu-id="c84a1-139">возможности пользователя по управлению вставкой содержимого.</span><span class="sxs-lookup"><span data-stu-id="c84a1-139">How much control the user has over any content insertion flows</span></span>

<span data-ttu-id="c84a1-140">Дополнительные сведения см. в статье [Эффективный выбор предметов исследования](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions).</span><span class="sxs-lookup"><span data-stu-id="c84a1-140">For more information, see [Writing Effective Questions.](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions)</span></span>
 
## <a name="3-identify-participants-to-target"></a><span data-ttu-id="c84a1-141">3. Определите участников тестирования</span><span class="sxs-lookup"><span data-stu-id="c84a1-141">3. Identify participants to target</span></span>
 
<span data-ttu-id="c84a1-p107">Удаленные службы тестирования позволяют контролировать множество характеристик участников тестирования. Задумайтесь, для каких пользователей предназначено исследование. На ранних этапах сбора данных рекомендуем набрать широкую аудиторию участников, чтобы выявить наиболее очевидные проблемы при использовании. Позже вам могут потребоваться конкретные целевые группы, например опытные пользователи Office, определенные профессии или возрастные категории.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p107">Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.</span></span>
 
## <a name="4-create-the-participant-screener"></a><span data-ttu-id="c84a1-146">4. Создайте анкету для отбора участников</span><span class="sxs-lookup"><span data-stu-id="c84a1-146">4. Create the participant screener</span></span>
 
<span data-ttu-id="c84a1-p108">Критерии отбора — это набор вопросов и требований, которые вы представляете потенциальным участникам тестирования. Помните, что участники в таких службах, как UserTesting.com, материально заинтересованы участвовать в вашем тестировании. Если вы хотите исключить участие определенных пользователей, рекомендуем задать несколько каверзных вопросов.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p108">The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test.</span></span> 
 
<span data-ttu-id="c84a1-150">Например, если вы хотите найти участников, знакомых с сайтом GitHub, добавьте в список вариантов ответа вымышленные названия, чтобы отфильтровать пользователей, предоставляющих ложные сведения.</span><span class="sxs-lookup"><span data-stu-id="c84a1-150">For example, if you want to find participants who are familiar with GitHub, to filter out users who might misrepresent themselves, include fakes in the list of possible answers.</span></span>

<span data-ttu-id="c84a1-151">**С каким из перечисленных ниже репозиториев кода вы знакомы?**</span><span class="sxs-lookup"><span data-stu-id="c84a1-151">**Which of the following source code repositories are you familiar with?**</span></span>  
 <span data-ttu-id="c84a1-p109">a. SourceShelf [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p109">a. SourceShelf  [*Reject*]</span></span>  
 <span data-ttu-id="c84a1-p110">b. CodeContainer [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p110">b. CodeContainer  [*Reject*]</span></span>  
 <span data-ttu-id="c84a1-p111">c. GitHub [*Должен выбрать*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p111">c. GitHub  [*Must select*]</span></span>  
 <span data-ttu-id="c84a1-p112">d. BitBucket [*Может выбрать*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p112">d. BitBucket  [*May select*]</span></span>  
 <span data-ttu-id="c84a1-p113">e. CloudForge [*Может выбрать*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p113">e. CloudForge  [*May select*]</span></span>  

<span data-ttu-id="c84a1-162">Если вы планируете тестировать готовую сборку надстройки, представленные ниже вопросы помогут вам отобрать пользователей, которые смогут это сделать.</span><span class="sxs-lookup"><span data-stu-id="c84a1-162">If you are planning to test a live build of your add-in, the following questions can screen for users who will be able to do this.</span></span> 

<span data-ttu-id="c84a1-163">**Для этого теста вам потребуется Microsoft PowerPoint 2016. У вас есть PowerPoint 2016?**</span><span class="sxs-lookup"><span data-stu-id="c84a1-163">**This test requires you to have Microsoft PowerPoint 2016. Do you have PowerPoint 2016?**</span></span>  
 <span data-ttu-id="c84a1-p114">a. Да [*Должен выбрать*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p114">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="c84a1-p115">b. Нет [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p115">b. No [*Reject*]</span></span>  
 <span data-ttu-id="c84a1-p116">c. Не знаю [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p116">c. I don’t know [*Reject*]</span></span>  

<span data-ttu-id="c84a1-170">**Для этого теста вам потребуется установить бесплатную надстройку для PowerPoint 2016 и создать бесплатную учетную запись для ее использования. Вы готовы установить надстройку и создать бесплатную учетную запись?**</span><span class="sxs-lookup"><span data-stu-id="c84a1-170">**This test requires you to install a free add-in for PowerPoint 2016, and create a free account to use it. Are you willing to install an add-in and create a free account?**</span></span>  
 <span data-ttu-id="c84a1-p117">a. Да [*Должен выбрать*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p117">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="c84a1-p118">b. Нет [*Отклонение*]</span><span class="sxs-lookup"><span data-stu-id="c84a1-p118">b. No [*Reject*]</span></span>  

<span data-ttu-id="c84a1-175">Дополнительные сведения см. в статье [Рекомендации по составлению вопросов для отбора](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices).</span><span class="sxs-lookup"><span data-stu-id="c84a1-175">For more information, see [Screener Questions Best Practices.](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices)</span></span>
 
## <a name="5-create-tasks-and-questions-for-participants"></a><span data-ttu-id="c84a1-176">5. Составьте список задач и вопросов для участников</span><span class="sxs-lookup"><span data-stu-id="c84a1-176">5. Create tasks and questions for participants</span></span>
 
<span data-ttu-id="c84a1-p119">Постарайтесь расставить приоритеты тестирования, чтобы ограничить количество задач и вопросов для каждого участника. Некоторые службы платят участникам только за определенное время, поэтому выполнение ваших заданий не должно занимать слишком много времени.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p119">Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.</span></span>

<span data-ttu-id="c84a1-p120">По возможности старайтесь наблюдать за действиями участников, а не спрашивать о них. Если вам нужно задавать вопросы о работе пользователей, спросите, что участники делали в таких ситуациях раньше, а не об их предположениях. Как правило, это дает более надежные результаты.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p120">Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.</span></span>
 
<span data-ttu-id="c84a1-p121">Самая сложная задача при тестировании без наблюдения — убедиться, что участники понимают ваши задания и контекст работы. Указания должны быть *понятными и краткими*. В неоднозначных указаниях кто-то обязательно запутается.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p121">The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.</span></span> 

<span data-ttu-id="c84a1-p122">Не полагайтесь на то, что у пользователя будет открыт именно тот экран, который необходим на этом этапе тестирования. Рекомендуем сообщать участникам, какой экран должен быть открыт в начале следующей задачи.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p122">Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.</span></span> 

<span data-ttu-id="c84a1-187">Дополнительные сведения см. в статье [Составление хороших задач](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks).</span><span class="sxs-lookup"><span data-stu-id="c84a1-187">For more information, see [Writing Great Tasks.](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks)</span></span>

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a><span data-ttu-id="c84a1-188">6. Создайте прототип, соответствующий задачам и вопросам</span><span class="sxs-lookup"><span data-stu-id="c84a1-188">6. Create a prototype to match the tasks and questions</span></span>
 
<span data-ttu-id="c84a1-p123">Вы можете тестировать либо готовую версию надстройки, либо ее прототип. Помните, что если вы хотите протестировать готовую версию надстройки, вам потребуется отобрать участников, у которых есть Office 2016 и которые готовы установить надстройку и зарегистрировать учетную запись (если у вас нет для них готовых учетных данных для входа). Затем необходимо убедиться, что они успешно установили надстройку.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p123">You can either test your live add-in, or you can test a prototype. Keep in mind that if you want to test the live add-in, you need to screen for participants that have Office 2016, are willing to install the add-in, and are willing to sign up for an account (unless you have logon credentials to provide them.) You'll then need to make sure that they successfully install your add-in.</span></span> 

<span data-ttu-id="c84a1-p124">В среднем обучение пользователей установке надстройки занимает около 5 минут. Ниже представлен пример понятных, кратких указаний по установке. Откорректируйте их в соответствии с особенностями вашего тестирования.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p124">On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.</span></span>

<span data-ttu-id="c84a1-194">**Установите надстройку <название надстройки> для PowerPoint 2016, выполнив следующие действия:**</span><span class="sxs-lookup"><span data-stu-id="c84a1-194">**Please install the (insert your add-in name here) add-in for PowerPoint 2016, using the following instructions:**</span></span> 

1. <span data-ttu-id="c84a1-195">Откройте Microsoft PowerPoint 2016.</span><span class="sxs-lookup"><span data-stu-id="c84a1-195">Open Microsoft PowerPoint 2016.</span></span>
2. <span data-ttu-id="c84a1-196">Выберите элемент **Пустая презентация**.</span><span class="sxs-lookup"><span data-stu-id="c84a1-196">Select **Blank Presentation.**</span></span>
3. <span data-ttu-id="c84a1-197">Нажмите **Вставка > Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="c84a1-197">Go to **Insert > My Add-ins.**</span></span>
5. <span data-ttu-id="c84a1-198">Во всплывающем окне выберите **Магазин**.</span><span class="sxs-lookup"><span data-stu-id="c84a1-198">In the popup window, choose **Store.**</span></span>
6. <span data-ttu-id="c84a1-199">Введите <название надстройки> в поле поиска.</span><span class="sxs-lookup"><span data-stu-id="c84a1-199">Type (Add-in name) in the search box.</span></span>
7. <span data-ttu-id="c84a1-200">Выберите <название надстройки>.</span><span class="sxs-lookup"><span data-stu-id="c84a1-200">Choose (Add-in name).</span></span>
8. <span data-ttu-id="c84a1-201">Изучите страницу Магазина, чтобы ознакомиться с надстройкой.</span><span class="sxs-lookup"><span data-stu-id="c84a1-201">Take a moment to look at the Store page to familiarize yourself with the add-in.</span></span>
9. <span data-ttu-id="c84a1-202">Нажмите **Добавить**, чтобы установить надстройку.</span><span class="sxs-lookup"><span data-stu-id="c84a1-202">Choose **Add** to install the add-in.</span></span>

<span data-ttu-id="c84a1-p125">Вы можете тестировать прототип на любом уровне взаимодействия и визуальной четкости. Для создания более сложных связей и повышенной интерактивности можно использовать такие средства создания прототипов, как [InVision](https://www.invisionapp.com). Если требуется протестировать статические экраны, вы можете разместить изображения в Интернете и отправить участникам соответствующий URL-адрес либо предоставить им ссылку на презентацию PowerPoint в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p125">You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation.</span></span> 

## <a name="7-run-a-pilot-test"></a><span data-ttu-id="c84a1-206">7. Проведите пилотный тест</span><span class="sxs-lookup"><span data-stu-id="c84a1-206">7. Run a pilot test</span></span>

<span data-ttu-id="c84a1-p126">Создание прототипа и составление списка задач или вопросов могут быть непростыми задачами. Пользователи могут запутаться в заданиях или не разобраться в прототипе. Следует провести пилотный тест с 1–3 пользователями, чтобы выявить неизбежные проблемы с форматом тестирования. Это поможет гарантировать, что ваши вопросы ясны, прототип настроен правильно и вы собираете именно те данные, которые вам нужны.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p126">It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.</span></span>

## <a name="8-run-the-test"></a><span data-ttu-id="c84a1-211">8. Проведите тестирование</span><span class="sxs-lookup"><span data-stu-id="c84a1-211">8. Run the test</span></span>

<span data-ttu-id="c84a1-p127">Заказав тест, вы будете получать по электронной почте уведомления, когда участники проходят его. Как правило, если у вас нет специфических требований к участникам, тесты завершаются в течение нескольких часов.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p127">After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.</span></span>

## <a name="9-analyze-results"></a><span data-ttu-id="c84a1-214">9. Проанализируйте результаты</span><span class="sxs-lookup"><span data-stu-id="c84a1-214">9. Analyze results</span></span>

<span data-ttu-id="c84a1-p128">На этом этапе вам необходимо сделать выводы из собранных данных. Просматривая видео тестирования, делайте заметки о проблемах и успехах пользователя. Старайтесь не делать поспешных выводов. Просмотрите все результаты.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p128">This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.</span></span> 

<span data-ttu-id="c84a1-p129">Проблема, возникшая у одного участника, еще не означает, что необходимо что-то менять. Если же одна и та же проблема возникнет у нескольких участников, значит, она возникнет и у других пользователей в широкой аудитории.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p129">A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.</span></span>

<span data-ttu-id="c84a1-p130">В целом, будьте осторожны с выводами при анализе данных. Не совершайте ошибку, стараясь подогнать результаты под определенную интерпретацию. Объективно оцените, что полученные данные доказывают или опровергают, а для каких выводов их просто недостаточно. Избегайте предвзятости — поведение пользователей часто противоречит ожиданиям разработчика.</span><span class="sxs-lookup"><span data-stu-id="c84a1-p130">In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.</span></span>
 

## <a name="see-also"></a><span data-ttu-id="c84a1-223">См. также</span><span class="sxs-lookup"><span data-stu-id="c84a1-223">See also</span></span>
 
 - [<span data-ttu-id="c84a1-224">Как проводить тестирование удобства использования</span><span class="sxs-lookup"><span data-stu-id="c84a1-224">How to Conduct Usability Testing</span></span>](http://whatpixel.com/howto-conduct-usability-testing/)  
 - [<span data-ttu-id="c84a1-225">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="c84a1-225">Best Practices</span></span>](http://help.usertesting.com/customer/en/portal/articles/1680726-best-practices)  
 - [<span data-ttu-id="c84a1-226">Предотвращение предвзятости</span><span class="sxs-lookup"><span data-stu-id="c84a1-226">Minimizing Bias</span></span>](http://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
