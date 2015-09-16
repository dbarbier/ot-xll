
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

#include <OT.hxx>

int WINAPI xlAutoOpen(void);
int WINAPI xlAutoClose(void);
int WINAPI xlAutoAdd(void);
int WINAPI xlAutoRemove(void);
void WINAPI xlAutoFree12(LPXLOPER12 p);
LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName);
LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction);

#define rgWorksheetFuncsRows 1
#define rgWorksheetFuncsCols 12

// Used To register XLL functions
LPWSTR rgWorksheetFuncs[rgWorksheetFuncsRows][rgWorksheetFuncsCols] =
{
    { L"OT_NORMAL_PDF",      // Name of the function in DLL
      L"UUUU",               // Data type of the return value and arguments
        // Most common values are B (double, passed by value) and
        // U (XLOPER12 values, arrays, and range references
      L"OT_NORMAL_PDF",      // The function name as it will appear in the Function Wizard
      L"Mu, Sigma, Point",   // Description of arguments
      L"1",                  // Macro type, use "1" by default or "2" for hidden commands
      L"Openturns Add-In",   // Category name
      L"",                   // Shortcut for commands, do not use
      L"",                   // Reference to the Help file
      L"Compute the probability density function",        // Function help
      L"Mean of the Gaussian distribution",               // Description of first argument
      L"Standard deviation of the Gaussian distribution", // Description of second argument
      L"Point where PDF is evaluated"                     // Description of third argument
    }
};


/******************************************************************************
** xlAutoOpen()
**
** Purpose:
**      Microsoft Excel call this function when the DLL is loaded.
**
**      Microsoft Excel uses xlAutoOpen to load XLL files.
**      When you open an XLL file, the only action
**      Microsoft Excel takes is to call the xlAutoOpen function.
**
**      More specifically, xlAutoOpen is called:
**
**       - when you open this XLL file from the File menu,
**       - when this XLL is in the XLSTART directory, and is
**         automatically opened when Microsoft Excel starts,
**       - when Microsoft Excel opens this XLL for any other reason, or
**       - when a macro calls REGISTER(), with only one argument, which is the
**         name of this XLL.
**
**      xlAutoOpen is also called by the Add-in Manager when you add this XLL
**      as an add-in. The Add-in Manager first calls xlAutoAdd, then calls
**      REGISTER("ot_simple_example.xll"), which in turn calls xlAutoOpen.
**
**      xlAutoOpen should:
**
**       - register all the functions you want to make available while this
**         XLL is open,
**
**       - add any menus or menu items that this XLL supports,
**
**       - perform any other initialization you need, and
**
**       - return 1 if successful, or return 0 if your XLL cannot be opened.
**
** Parameters:
**
** Returns:
**
**      int         1 on success, 0 on failure
*****************************************************************************/
int WINAPI xlAutoOpen(void)
{

    static XLOPER12 xDLL; // name of this DLL
    int i;                // Loop index

    /*
    ** In the following block of code the name of the XLL is obtained by
    ** calling xlGetName. This name is used as the first argument to the
    ** REGISTER function to specify the name of the XLL. Next, the XLL loops
    ** through the rgFuncs[] table, registering each function in the table using
    ** xlfRegister. Functions must be registered before you can add a menu
    ** item.
    */

    Excel12f(xlGetName, &xDLL, 0);

    for (i=0;i<rgWorksheetFuncsRows;i++)
    {
        Excel12f(xlfRegister, 0,  1 + rgWorksheetFuncsCols,
            (LPXLOPER12)&xDLL,
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][0]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][1]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][2]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][3]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][4]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][5]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][6]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][7]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][8]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][9]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][10]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][11])
            );
    }

    /* Free the XLL filename */
    Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);
    return 1;
}

/*********************************************************************************
** xlAutoRegister12()
**
** Purpose:
**
**      This function is called by Microsoft Excel if a macro sheet tries to
**      register a function without specifying the type_text argument. If that
**      happens, Microsoft Excel calls xlAutoRegister12, passing the name of the
**      function that the user tried to register. xlAutoRegister12 should use the
**      normal REGISTER function to register the function, only this time it must
**      specify the type_text argument. If xlAutoRegister12 does not recognize the
**      function name, it should return a #VALUE! error. Otherwise, it should
**      return whatever REGISTER returned.
**
** Parameters:
**
**      LPXLOPER12 pxName   xltypeStr containing the
**                          name of the function
**                          to be registered. This is not
**                          case sensitive.
**
** Returns:
**
**      LPXLOPER12          xltypeNum containing the result
**                          of registering the function,
**                          or xltypeErr containing #VALUE!
**                          if the function could not be
**                          registered.
***********************************************************************************/
LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
    static XLOPER12 xDLL, xRegId;
    int i;
    xRegId.xltype = xltypeErr;
    xRegId.val.err = xlerrValue;

    for (i = 0; i < rgWorksheetFuncsRows; i++)
    {
        if (_wcsicmp(rgWorksheetFuncs[i][0], pxName->val.str)==0)
        {
            Excel12f(xlfRegister, 0, 1 + rgWorksheetFuncsCols,
                    (LPXLOPER12) &xDLL,
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][0]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][1]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][2]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][3]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][4]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][5]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][6]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][7]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][8]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][9]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][10]),
                    (LPXLOPER12) TempStr12(rgWorksheetFuncs[i][11])
                    );

            // Free the oper returned by Excel.
            Excel12f(xlFree, 0, 1, (LPXLOPER12) &xDLL);
            return(LPXLOPER12) &xRegId;
        }
    }

    return(LPXLOPER12) &xRegId;
}


/************************************************************************************
** xlAutoClose
**
** xlAutoClose is called by Microsoft Excel:
**
**  - when you quit Microsoft Excel, or
**  - when a macro sheet calls UNREGISTER(), giving a string argument
**        which is the name of this XLL.
**
** xlAutoClose is called by the Add-in Manager when you remove this XLL from
** the list of loaded add-ins. The Add-in Manager first calls xlAutoRemove,
** then calls UNREGISTER("ot_simple_example"), which in turn calls xlAutoClose.
**
**
** xlAutoClose should:
**
**  - Remove any menus or menu items that were added in xlAutoOpen,
**
**  - do any necessary global cleanup, and
**
**  - delete any names that were added (names of exported functions, and
**        so on). Remember that registering functions may cause names to be created.
**
** xlAutoClose does NOT have to unregister the functions that were registered
** in xlAutoOpen. This is done automatically by Microsoft Excel after
** xlAutoClose returns.
**
** xlAutoClose should return 1.
***********************************************************************************/
int WINAPI xlAutoClose(void)
{
    int i;

    /*
    ** This block first deletes all names added by xlAutoOpen or by
    ** xlAutoRegister.
    */

    for (i = 0; i < rgWorksheetFuncsRows; i++)
        Excel12f(xlfSetName, 0, 1, TempStr12(rgWorksheetFuncs[i][2]));
    return 1;
}


/**************************************************************************
** xlAutoAdd
**
** This function is called by the Add-in Manager only. When you add a
** DLL to the list of active add-ins, the Add-in Manager calls xlAutoAdd()
** and then opens the XLL, which in turn calls xlAutoOpen.
**
***************************************************************************/
int WINAPI xlAutoAdd(void)
{
    XCHAR szBuf[255];

    wsprintfW((LPWSTR)szBuf, L"Thank you for adding ot_simple_example.xll\n"
                             L" build date %hs, time %hs, OpenTURNS %hs",__DATE__, __TIME__,OT::PlatformInfo::GetVersion().c_str());

    /* Display a dialog box indicating that the XLL was successfully added */
    Excel12f(xlcAlert, 0, 2, TempStr12(szBuf), TempInt12(2));
    return 1;
}



/**************************************************************************
** xlAutoRemove
**
** This function is called by the Add-in Manager only. When you remove
** an XLL from the list of active add-ins, the Add-in Manager calls
** xlAutoRemove() and then UNREGISTER("ot_simple_example").
**
** You can use this function to perform any special tasks that need to be
** performed when you remove the XLL from the Add-in Manager's list
** of active add-ins. For example, you may want to delete an
** initialization file when the XLL is removed from the list.
***************************************************************************/
int WINAPI xlAutoRemove(void)
{
    /* Display a dialog box indicating that the XLL was successfully removed */
    Excel12f(xlcAlert, 0, 2, TempStr12(L"Thank you for removing ot_simple_example.xll!"), TempInt12(2));
    return 1;
}


/******************************************************************************
** xlAddInManagerInfo12()
**
** Purpose:
**
**      This function is called by the Add-in Manager to find the long name
**      of the add-in. If xAction = 1, this function should return a string
**      containing the long name of this XLL, which the Add-in Manager will use
**      to describe this XLL. If xAction = 2 or 3, this function should return
**      #VALUE!.
**
** Parameters:
**
**      LPXLOPER12 xAction  What information you want. One of:
**                            1 = the long name of the
**                                add-in
**                            2 = reserved
**                            3 = reserved
**
** Returns:
**
**      LPXLOPER12          The long name or #VALUE!.
******************************************************************************/
LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction)
{
    static XLOPER12 xInfo, xIntAction;

#ifdef _DEBUG
    debugPrintf("xlAutoAddInManagerInfo12\n");
#endif

    // This code coerces the passed-in value to an integer.
    Excel12f(xlCoerce, &xIntAction, 2, xAction, TempInt12(xltypeInt));

    if (xIntAction.val.w == 1)
    {
        // Note that the string is length-prefixed in octal.
        xInfo.xltype = xltypeStr;
        xInfo.val.str = L"\030OpenTurns Standalone DLL";
    }
    else
    {
        xInfo.xltype = xltypeErr;
        xInfo.val.err = xlerrValue;
    }

    // Word of caution: returning static XLOPERs/XLOPER12s is
    // not thread-safe. For UDFs declared as thread-safe, use
    // alternate memory allocation mechanisms.
    return(LPXLOPER12) &xInfo;
}



/*************************************************************************
** xlAutoFree
**
** Frees the memory allocated.
** Called by Microsoft Excel just after an XLL worksheet function returns
** an XLOPER/XLOPER12 to it with a flag set that tells it there is memory
** that the XLL still needs to release. This enables the XLL to return
** dynamically allocated arrays, strings, and external references
** to the worksheet without memory leaks
**
*************************************************************************/
void WINAPI xlAutoFree12(LPXLOPER12 p)
{
    if (p->xltype == (xltypeMulti | xlbitDLLFree))
    {
        int size = p->val.array.rows * p->val.array.columns;
        LPXLOPER12 parray = p->val.array.lparray;

        for(; size-- > 0; parray++) // check elements for strings
            if(parray->xltype == xltypeStr)
                delete [] parray->val.str;
        delete [] p->val.array.lparray;
    }
    else if(p->xltype  == ( xltypeStr | xlbitDLLFree))
    {
        delete [] p->val.str;
    }
    else if(p->xltype == ( xltypeRef | xlbitDLLFree))
    {
        delete [] p->val.mref.lpmref;
    }
    // Assume p was itself dynamically allocated using new.
    delete p;
}
