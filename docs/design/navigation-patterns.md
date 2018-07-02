# <a name="navigation-patterns"></a><span data-ttu-id="02817-101"> Шаблоны навигации</span><span class="sxs-lookup"><span data-stu-id="02817-101">Navigation patterns</span></span>

<span data-ttu-id="02817-102">Доступ к основным функциям надстройки осуществляется через определенные типы команд и ограниченную площадь экрана.</span><span class="sxs-lookup"><span data-stu-id="02817-102">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="02817-103">Важно, что навигация интуитивно понятна, обеспечивает контекст и позволяет пользователю легко перемещаться по всей надстройке.</span><span class="sxs-lookup"><span data-stu-id="02817-103">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="02817-104">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="02817-104">Best practices</span></span>

| <span data-ttu-id="02817-105">Правильно</span><span class="sxs-lookup"><span data-stu-id="02817-105">Do</span></span>    | <span data-ttu-id="02817-106">Неправильно</span><span class="sxs-lookup"><span data-stu-id="02817-106">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="02817-107">Убедитесь, что у пользователя есть видимая опция навигации.</span><span class="sxs-lookup"><span data-stu-id="02817-107">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="02817-108">Не затрудняйте процесс навигации, используя нестандартный пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="02817-108">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="02817-109">Используйте, по возможности, следующие компоненты, чтобы пользователи могли перемещаться по вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="02817-109">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="02817-110">Не затрудняйте понимание пользователем своего текущего места или контекста в надстройке</span><span class="sxs-lookup"><span data-stu-id="02817-110">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="02817-111">Панель команд</span><span class="sxs-lookup"><span data-stu-id="02817-111">command bar</span></span>

<span data-ttu-id="02817-112">CommandBar - это поверхность, на которой размещаются команды, которые работают с содержимым окна, панели или родительской области, в которой она находится выше.</span><span class="sxs-lookup"><span data-stu-id="02817-112">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="02817-113">Дополнительные функции включают точку доступа к меню "гамбургера", поиск и боковые команды.</span><span class="sxs-lookup"><span data-stu-id="02817-113">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Команды - это спецификации для панели задач на рабочем столе](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="02817-115">Панель вкладок</span><span class="sxs-lookup"><span data-stu-id="02817-115">Tab bar</span></span>

<span data-ttu-id="02817-116">Показывает панель навигации, используя кнопки с расположенными по вертикали в столбик текстом и значками.</span><span class="sxs-lookup"><span data-stu-id="02817-116">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="02817-117">Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.</span><span class="sxs-lookup"><span data-stu-id="02817-117">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Панель вкладок - это технические характеристики панели задач на рабочем столе](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="02817-119">Кнопка "Назад"</span><span class="sxs-lookup"><span data-stu-id="02817-119">Back button</span></span>

<span data-ttu-id="02817-120">Кнопка «Назад» позволяет пользователям восстанавливаться после детализированного навигационного действия.</span><span class="sxs-lookup"><span data-stu-id="02817-120">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="02817-121">Этот шаблон помогает пользователям следовать упорядоченной последовательности шагов.</span><span class="sxs-lookup"><span data-stu-id="02817-121">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![Кнопка «Назад» - это спецификация для панели задач на рабочем столе](../images/add-in-back-button.png)
