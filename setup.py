{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 6,
   "metadata": {},
   "outputs": [
    {
     "ename": "ModuleNotFoundError",
     "evalue": "No module named 'cx_Freeze'",
     "output_type": "error",
     "traceback": [
      "\u001b[1;31m---------------------------------------------------------------------------\u001b[0m",
      "\u001b[1;31mModuleNotFoundError\u001b[0m                       Traceback (most recent call last)",
      "\u001b[1;32m<ipython-input-6-5f6458dd3288>\u001b[0m in \u001b[0;36m<module>\u001b[1;34m\u001b[0m\n\u001b[1;32m----> 1\u001b[1;33m \u001b[1;32mfrom\u001b[0m \u001b[0mcx_Freeze\u001b[0m \u001b[1;32mimport\u001b[0m \u001b[0msetup\u001b[0m\u001b[1;33m,\u001b[0m \u001b[0mExecutable\u001b[0m\u001b[1;33m\u001b[0m\u001b[1;33m\u001b[0m\u001b[0m\n\u001b[0m\u001b[0;32m      2\u001b[0m \u001b[1;33m\u001b[0m\u001b[0m\n\u001b[0;32m      3\u001b[0m setup(name = \"reandurllib\" ,\n\u001b[0;32m      4\u001b[0m       \u001b[0mversion\u001b[0m \u001b[1;33m=\u001b[0m \u001b[1;34m\"0.1\"\u001b[0m \u001b[1;33m,\u001b[0m\u001b[1;33m\u001b[0m\u001b[1;33m\u001b[0m\u001b[0m\n\u001b[0;32m      5\u001b[0m       \u001b[0mdescription\u001b[0m \u001b[1;33m=\u001b[0m \u001b[1;34m\"\"\u001b[0m \u001b[1;33m,\u001b[0m\u001b[1;33m\u001b[0m\u001b[1;33m\u001b[0m\u001b[0m\n",
      "\u001b[1;31mModuleNotFoundError\u001b[0m: No module named 'cx_Freeze'"
     ]
    }
   ],
   "source": [
    "from cx_Freeze import setup, Executable\n",
    "\n",
    "setup(name = \"reandurllib\" ,\n",
    "      version = \"0.1\" ,\n",
    "      description = \"\" ,\n",
    "      executables = [Executable(\"First.py\")])"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.7.4"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 2
}