<!--#include virtual="\_work\bugs\_private\config.asp"-->
<%If bSuperAdminAccess Then%>
	<h1>NewZapp Dev List</h1>
    <ol>
      <li><a href="http://orderentry-soap.destinet.co.uk/_work/bugs/viewList.asp?workID=&workType=0&viewType=1&sortType=0&projectID=0&priorityOperator=<=&priorityLevel=5&bugSection=">FULL NEWZAPP LIST</a><br /></li>
	  
	  <li><a href="viewWishList.asp">Wish List</a><br />
            </li>
      <li><a href="viewBugList.asp">Bug List</a><br />
            </li>
      <li><a href="viewTaskList.asp">Task List</a><br /></li>
	  
	  
    </ol>
        <%Else%>
        Access Denied
  <%End If%>
                      
    