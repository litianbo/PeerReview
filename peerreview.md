This is a step-by-step instruction on how to run the peer review program.
===
Tianbo Li
---
tianboli@usc.edu
---
### Before you run the program: ###
1. Download two libraries: poi-bin-3.11-20141221 and java_ee_sdk-7u1 
2. Make appropriate reference in the build path so that the program could find these two libraries.
3. Read the following two turtorials:　
   * http://howtodoinjava.com/2013/06/19/readingwriting-excel-files-in-java-poi-tutorial/
   * http://www.tutorialspoint.com/java/java_sending_email.htm
4. Register a new email address so that you could use it to send emails to everyone.
    ### Algorithm ###
* The idea is very simply, use arrays to store data from excel file and output them into another excel file in order of their IDs.
* Then, one can use javax library to send email to everyone.
* One thing to notice: the data is not well-structured. So it requires careful mapping between the names on the survey and names on the blackboard.
