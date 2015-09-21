#ifndef __XLL_HELPER_FUNCTIONS_H
#define __XLL_HELPER_FUNCTIONS_H

#ifndef WIN32_LEAN_AND_MEAN
# define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>
#include <string>

/* Coerce an XLOPER12 to a */
int xloper_to_multi(LPXLOPER12 p_op, LPXLOPER12 ret_val);
/* Convert an XLOPER12 to a double  */
int xloper_to_num(LPXLOPER12 xl_poper, double* value);
/* Convert an XLOPER12 to an int  */
int xloper_to_int(LPXLOPER12 xl_poper, int* value);

/* Get the number of rows of active selection */
int getNumberOfRows();
/* Get the number of columns of active selection */
int getNumberOfColumns();

/* Display an error message */
LPXLOPER12 dialogError(const std::string & msg, int error_code);

/* Check whether function is called from Function Wizard */
bool isCalledByFuncWiz();

#endif // __XLL_HELPER_FUNCTIONS_H

