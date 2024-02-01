# GhostWriter

Script takes a series of questions with related answare from a docx file, put then in a SQLite db and run in an infinite loop.
When script is running, user can selects with mouse cursor a text on the screen, after copy the text with control+c in the clipboard, pressing altgr and script starting a match between selected test and questions in db to find related best answer, so starting writing the answer where user pointing with cursor.

USES CASE: adding into forlder a docx with question-answer format. Question must be start with a number (1 to 9) and number cannot be indentated. Document must not use any tipe of intendatition

There are 2 versions
1) controlc.py: behaviours is the same descripted before
2) controlcstopspace.py: adding stop writing when space is pressed

Other test version is
3) multiprocess.py: tries to implement restore writing after stop with multiprocessing, but doesn't work

Limitations
- controlc.py cannot be stopped after start writing and can be speed up or slow down based on interval value
- controlcstopspace.py can be stopped but after stop, there are no ways to restart the writing from point where it had been stopped
- controlcstopspace.py speeds of writing cannot be modified because of character loop, interval value doesn't work for single character speed
- multiprocess.py doesn't work
