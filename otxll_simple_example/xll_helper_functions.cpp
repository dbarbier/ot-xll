//                                               -*- C++ -*-
/**
 *  Copyright 2005-2015 Airbus-IMACS
 *
 *  This library is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU Lesser General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  along with this library.  If not, see <http://www.gnu.org/licenses/>.
 *
 */
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>
#include <iostream>

#include "xll_helper_functions.h"

/****************************************************************
** xloper_to_multi()
**
** Purpose:
**
**      This function takes 2 argument, coerces xloper to 2D array
**
** Parameters:
**
**      LPXLOPER12      2 argument : p_op, ret_val
**
** Returns:
**
**      LPXLOPER12      xloper of xltypeMulti Type.
*****************************************************************/
int xloper_to_multi(LPXLOPER12 p_op,
                    LPXLOPER12 ret_val)
{
    int error = -1;
    switch (p_op->xltype)
    {
    case xltypeNum:
        break;
    case xltypeStr:
        break;
    case xltypeRef:
    case xltypeSRef:
    case xltypeMulti:
        if ( xlretSuccess  != Excel12f( xlCoerce,
                                        ret_val,
                                        2,
                                        p_op,
                                        TempInt12(xltypeMulti)))
        {
            return 0;
        }
        break;
    case xltypeMissing:
        break;
    case xltypeNil:
        break;
    case xltypeErr:
        error = p_op->val.err;
        break;
    default:
        error = xlerrValue;
        break;
    }

    return error;
}


/*********************************************************************
 xloper_to_num()

 Purpose:

      This function takes 2 argument, coerces xloper to numerical
      type and gets numerical value.

 Parameters:

      LPXLOPER12  xl_poper: Excel cell
      double *    value: pointer to a double, where result is stored

 Returns:

      int: -1 if conversion was successful, 0 otherwise.

************************************************************************/

int
xloper_to_num(LPXLOPER12 xl_poper, double* value)
{
    XLOPER12 xl_oper;
    int error = -1;
    int xlerror;

    //  Switch on XLOPER TYPE
    switch (xl_poper->xltype)
    {

    // xloper is of numerical type
    case xltypeNum:
        *value = xl_poper->val.num;
        break;
    // excel reference to a cell (current  or not current sheet )
    case xltypeRef:
    case xltypeSRef:
        // Excel12f is a robust wrapper for the Excel12() function
        // It also does the following:
        // (1) Checks that none of the LPXLOPER12 arguments are 0,
        //        which would indicate that creating a temporary XLOPER12
        //        has failed. In this case, it doesn't call Excel12
        //        but it does print a debug message.
        //  (2) If an error occurs while calling Excel12,
        //        print a useful debug message.
        //  (3) When done, free all temporary memory.
        xlerror = Excel12f(xlCoerce, &xl_oper, 2,
                           xl_poper, TempInt12(xltypeNum));
        if(xlerror != xlretSuccess)
        {
            error = xlerror;
        }
        *value = xl_oper.val.num;

        // Free the XLOPER12 returned by xlCoerce
        Excel12f(xlFree, 0, 1, (LPXLOPER12) &xl_oper);
        break;
    // excel type error
    case xltypeErr:
        error = xl_poper->val.err;
        break;
    // other cases
    default:
        error = xlerrValue;
        break;
    }

    return error;
}

/*******************************************************************
** xloper_to_int()
**
** Purpose:
**
**      This function takes 2 argument, coerces xloper to integer
**      type and gets numerical value.
**
** Parameters:
**
**      LPXLOPER12      2 argument : xl_poper, value
**      int          Integer value of xloper.
**
** Returns:
**
**      -1 if success, error else
******************************************************************/
int xloper_to_int(LPXLOPER12 xl_poper,
                  int* value)
{
    XLOPER12 xl_oper;
    int error = -1;
    int xlerror;


    switch (xl_poper->xltype)
    {
    case xltypeNum:
        *value = (int)xl_poper->val.num;
    case xltypeInt:
        *value = xl_poper->val.w;
    case xltypeRef:
    case xltypeSRef:
        xlerror = Excel12f( xlCoerce,
                            &xl_oper,
                            2,
                            xl_poper,
                            TempInt12(xltypeInt));

        if(xlerror != xlretSuccess)
        {
            error = xlerror;
        }
        *value = xl_oper.val.w;

        // Free the XLOPER12 returned by xlCoerce
        Excel12f(xlFree, 0, 1, (LPXLOPER12) &xl_oper);
        break;
    case xltypeErr:
        error = xl_poper->val.err;
        break;
    default:
        error = xlerrValue;
        break;
    }

    return error;
}

/*********************************************************************
**  getNumberOfRows()
**
**  Purpose :
**        returns the number of rows of selected cells.
**
**
**  Returns :
**      the number of rows of current selection, or 0 if there is an error
**********************************************************************/
int getNumberOfRows()
{
    XLOPER12 Caller;
    XLOPER12 ret_val;

    if(xlretSuccess != Excel12f(xlfCaller, &Caller, 0))
    {
        return 0;
    }

    if(xlretSuccess != Excel12f( xlCoerce,
            &ret_val,
            2,
            &Caller,
            TempInt12(xltypeMulti)))
    {
        return 0;
    }

    int result = ret_val.val.array.rows;

    // Free the XLOPER12 returned by xlCoerce && xlfCaller
    Excel12f(xlFree, 0, 1, &Caller);
    Excel12f(xlFree, 0, 1, &ret_val);

    return result;
}


/*********************************************************************
**  dialogError()
**
**  Purpose :
**      display a dialog message to report some error message.
**
**  Parameters:
**
**        msg : std::string
**              message
**
**  Returns :
**        LPXLOPER12 so that caller functions can use
**          return dialogError("Some message")
**********************************************************************/
LPXLOPER12
dialogError(const std::string & msg, int error_code)
{
    // Until we find a better way, print this message on console.
    // Most of the time there is no console and message won't get
    // displayed, but it looks like user defined functions cannot
    // display dialog boxes
    std::cout << msg << std::endl;

    LPXLOPER12 xResult = new XLOPER12();
    xResult->xltype = xltypeErr | xlbitDLLFree;
    xResult->val.err = error_code;
    return xResult;
}
