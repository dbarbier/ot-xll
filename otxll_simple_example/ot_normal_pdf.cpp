#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

#include <OT.hxx>

/*********************************************************************
 xloper_to_multi()

 Purpose:

      This function takes 2 argument, coerce xloper to numerical
      type and get numerical value.

 Parameters:

      LPXLOPER12  xl_poper: Excel cell
      double *    value: pointer to a double, where result is stored

 Returns:

      int: -1 if conversion was successful, 0 otherwise.

************************************************************************/

static int
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

    // Free the XLOPER12 returned by xlCoerce
    Excel12f(xlFree, 0, 1, (LPXLOPER12) &xl_oper);

    return error;
}

/***********************************************************************************
 OT_NORMAL_PDF()

 Purpose:

      This function takes 3 argument and computes the normal distribution at a point.

 Parameters:

      LPXLOPER12      3 argument : xl_mu, xl_sigma, xl_point
                      (can be reference or value)

 Returns:

      LPXLOPER12      the normal distribution at a point
                      or #VALUE! if there are
                      non-numerics in the supplied
                      argument.
*************************************************************************************/

LPXLOPER12 WINAPI
OT_NORMAL_PDF(LPXLOPER12 xl_mu, LPXLOPER12 xl_sigma, LPXLOPER12 xl_point)
{
    LPXLOPER12 xResult = new XLOPER12();
    double mu, sigma, point;
    int error = -1;

    // Coerce the mean parameter
    //======================
    if((error = xloper_to_num(xl_mu, &mu)) != -1)
    {
        xResult->xltype = xltypeErr | xlbitDLLFree;
        xResult->val.err = error;
        return xResult;
    }

    // Coerce the standard deviation parameter
    //========================================
    if((error = xloper_to_num(xl_sigma, &sigma)) != -1)
    {
        xResult->xltype = xltypeErr | xlbitDLLFree;
        xResult->val.err = error;
        return xResult;
    }

    // Coerce the point parameter : to compute the PDF
    //================================================
    if((error = xloper_to_num(xl_point, &point)) != -1)
    {
        xResult->xltype = xltypeErr | xlbitDLLFree;
        xResult->val.err = error;
        return xResult;
    }

    //Compute the PDF on point
    //========================
    try
    {
        OT::Normal distribution(mu, sigma);
        // xlbitDLLFree enables the DLL to release
        // any dynamically allocated memory
        // that was associated with the xloper
        xResult->xltype = xltypeNum | xlbitDLLFree;
        xResult->val.num = distribution.computePDF(point);
    }
    catch(OT::Exception & e)
    {
        delete xResult;
        return 0;
    }

    return xResult;
}

