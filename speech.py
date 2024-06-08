import speech_recognition as sr
import pyaudio
import tkinter as tk
import os
import aspose.words as aw
import nltk



def speak():

    r = sr.Recognizer()
    with sr.Microphone() as source:
       
        r.adjust_for_ambient_noise(source, duration=1)
        print("Listening...")
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=50)
            query = r.recognize_google(audio)
            print("You said: " + query)
            return query
        except sr.WaitTimeoutError:
            print("Listening timed out while waiting for phrase to start")
        






def start_gui():

    window=tk.Tk()
    window.geometry("1350x690+0+0")
    window.title("REALTIME AUDIO REPORT PROJECT")
    window.configure(bg="#856ff8")

    def start():

        def downloadreport(query,splitword,maxword,noun,adjective,adverb,verb):
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            builder.write("\n")
            builder.write(f"Original Sentence-: {query}")
            builder.write("\n")

            builder.write("\n")
            builder.write(f"Total Words-: {len(splitword)}")
            builder.write("\n")

            builder.write("\n")
            builder.write(f"Most Used Word-: {maxword}")
            builder.write("\n")

            builder.write("\n")
            builder.write(f"NOUNS-: {noun}")
            builder.write("\n")

            builder.write("\n")
            builder.write(f"ADJECTIVES-: {adjective}")
            builder.write("\n")

            builder.write("\n")
            builder.write(f"VERBS-: {verb}")
            builder.write("\n")

            builder.write("\n")
            builder.write(f"ADVERBS-: {adverb}")
            builder.write("\n")

            doc.save("speechreport.docx")
            os.startfile("speechreport.docx")

        def seereport(query,origsent,origsent1,sreport,splitword,maxword,noun,adjective,adverb,verb):
            sreport.destroy()
            origsent1.destroy()
            origsent.destroy()
            leng = tk.Label(text="Total words-: ", bg="#856ff8", font=("Times New Roman", 15), fg='white')
            leng.place(x=550, y=220)
            leng1 = tk.Label(text=len(splitword), bg="#856ff8", wraplength=500, font=("Times New Roman", 15),
                             fg='black')
            leng1.place(x=760, y=220)

            maxwor = tk.Label(text="Word that you have used the most-: ", bg="#856ff8", font=("Times New Roman", 15),
                              fg='white')
            maxwor.place(x=550, y=270)
            maxwor1 = tk.Label(text=maxword, bg="#856ff8", wraplength=500, font=("Times New Roman", 15), fg='black')
            maxwor1.place(x=900, y=270)

            nun = tk.Label(text="NOUNS-: ", bg="#856ff8", font=("Times New Roman", 15), fg='white')
            nun.place(x=550, y=320)
            nun1 = tk.Label(text=noun, bg="#856ff8", wraplength=500, font=("Times New Roman", 15),
                            fg='black')
            nun1.place(x=760, y=320)

            adj = tk.Label(text="ADJECTIVES-: ", bg="#856ff8", font=("Times New Roman", 15), fg='white')
            adj.place(x=550, y=370)
            adj1 = tk.Label(text=adjective, bg="#856ff8", wraplength=500, font=("Times New Roman", 15),
                            fg='black')
            adj1.place(x=760, y=370)

            adve = tk.Label(text="ADVERBS-: ", bg="#856ff8", font=("Times New Roman", 15), fg='white')
            adve.place(x=550, y=420)
            adve1 = tk.Label(text=adverb, bg="#856ff8", wraplength=500, font=("Times New Roman", 15),
                             fg='black')
            adve1.place(x=760, y=420)

            ver = tk.Label(text="VERBS-: ", bg="#856ff8", font=("Times New Roman", 15), fg='white')
            ver.place(x=550, y=470)
            ver1 = tk.Label(text=verb, bg="#856ff8", wraplength=500, font=("Times New Roman", 15),
                            fg='black')
            ver1.place(x=760, y=470)

            dreport = tk.Button(window, text="Download Report",
                                command=lambda: downloadreport(query,splitword, maxword, noun,
                                                          adjective, adverb, verb), bg='black',
                                font=('Helvetica 20 bold italic'),
                                fg='white')
            dreport.place(x=900, y=570)


        def generatereport(query):
            splitword = query.split(" ")
            worddict = {}
            for i in splitword:
                worddict[i] = 0

            for i in splitword:
                worddict[i] = worddict[i] + 1
            v = list(worddict.values())
            k = list(worddict.keys())
            maxword = k[v.index(max(v))]

            tagged = nltk.pos_tag(splitword)

            noun = []
            verb = []
            adverb = []
            adjective = []


            for i in range(len(tagged)):
                if (tagged[i][1][0] == 'N'):
                    noun.append(tagged[i][0])
                elif (tagged[i][1][0] == 'V'):
                    verb.append(tagged[i][0])
                elif (tagged[i][1][0] == 'J'):
                    adjective.append(tagged[i][0])
                elif (tagged[i][1][0] == 'R'):
                    adverb.append(tagged[i][0])


            print("-----------YOUR AUDIO REPORT---------")

            print("You have speaked-: ", query)
            print("Total words you have speaked-: ", len(splitword))
            print("The word you have used the most is-: ", maxword)
            print("NOUNS-: ", end=" ")
            for i in noun:
                print(i, end=" ")
            print()
            print("VERBS-: ", end=" ")
            for i in verb:
                print(i, end=" ")
            print()
            print("ADJECTIVE-: ", end=" ")
            for i in adjective:
                print(i, end=" ")
            print()
            print("ADVERB-: ", end=" ")
            for i in adverb:
                print(i, end=" ")
            print()


            headlbl = tk.Label(text="*YOUR AUDIO REPORT*", bg="#856ff8", font=("Times New Roman", 20), fg='white')
            headlbl.place(x=750, y=70)

            origsent = tk.Label(text="ORIGNAL SENTENCE-: ", bg="#856ff8", font=("Times New Roman", 15), fg='white')
            origsent.place(x=550, y=170)
            origsent1 = tk.Label(text=query, bg="#856ff8", wraplength=500, font=("Times New Roman", 15), fg='black')
            origsent1.place(x=760, y=170)

            sreport = tk.Button(window, text="See Report", command=lambda: seereport(query,origsent,origsent1,sreport,splitword,maxword,noun,adjective,adverb,verb), bg='black',
                             font=('Helvetica 20 bold italic'),
                             fg='white')
            sreport.place(x=900,y=570)










        def speakfunction():

            query1=speak()

            btn1 = tk.Button(window, text="Generate Report", command=lambda:generatereport(query1), bg='black', font=('Helvetica 20 bold italic'),
                             fg='white')
            btn1.place(x=170, y=470, width=250, height=80)



        photo = tk.PhotoImage(file=r"mic.png")
        image_lbl=tk.Label(image=photo)

        btn = tk.Button(window,image=photo,command=lambda:speakfunction(),border=0,borderwidth=0)

        btn.place(x=150,y=30)



        window.mainloop()

    start()



start_gui()