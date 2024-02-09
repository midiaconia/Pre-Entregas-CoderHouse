from flask import Flask, render_template, request
from pymongo import MongoClient
from dotenv import load_dotenv
import datetime
import os
import openai
from bson import ObjectId

load_dotenv()

def create_app():
    app = Flask(__name__)
    client = MongoClient(os.getenv("MONGODB_URI"))
    openai_api_key = os.getenv("OPENAI_API_KEY")
    openai.api_key = openai_api_key
    app.db = client.openai

    @app.route('/', methods=["GET", "POST"])
    def openai_route():
        entries_with_date = []
        entry_id = None  # Initialize entry_id to None at the top
        ideas_list = None  # Initialize with None or an empty string

        if request.method == "POST":
            ppt_topic = request.form.get("ppt_topic")
            ppt_slides = request.form.get("ppt_slides")
            try:
                ppt_slides = int(ppt_slides)
                if 1 <= ppt_slides <= 8:
                    # Generate the presentation content using OpenAI
                    conversation=[
                        {"role":"system","content":"Actúa como Rony Starp, un asistente diseñador experto en la generación "\
                        +"de diapositivas en Powerpoint, con más de 10 años de experiencia trabajando para uno de los mejores "\
                        +"expositores y oradores a nivel global. Ayúdame diseñando una presentación en PowerPoint, por lo que "\
                        +"debes proponer un título para cada diapositiva. Enfócate en armar "\
                        +"una presentación sobre " + ppt_topic + " con " + str(ppt_slides) + " diapositivas. Por favor "\
                        +"utiliza el símbolo '|' para separar los títulos de cada diapositiva, anteponiendo ' | ' cada vez "\
                        +"que uses la palabra Diapositiva (o su equivalente en el idioma en el que el usuario "\
                        +"ingresa" + ppt_topic + "sin contar la primera vez que usas dicha palabra, refiriéndote a la "\
                        +"Diapositiva 1. Utiliza este ejemplo: 'Diapositiva 1: [TÍTULO] | Diapositiva 2: [TÍTULO] | "\
                        +"Diapositiva 3: [TÍTULO] y así sucesivamente."\
                        +"Es indispensable que agregues el caracter '|' según tus indicaciones."\
                        +"Finalmente, empieza tu respuesta entregando defrente el título de la Diapositiva 1, sin saludo previo. "\
                        +"No me hagas preguntas sobre el diseño, ya que no es mi especialidad."},
                        {"role":"user","content":"Dame una lista de " + str(ppt_slides) + " títulos que representarán las ideas "\
                        +"centrales que formarán el índice de la presentación sobre " + ppt_topic + "."}
                    ]

                    response_presentation = openai.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=conversation,
                        max_tokens=1000,
                        temperature=0.7
                    )

                    # The first element may be empty or not relevant if "diapositiva" is at the start of every slide's content
                    ideas_list_str  = response_presentation.choices[0].message.content

                    # Splitting the response into a list of slides based on the "|" delimiter
                    ideas_list = ideas_list_str.split('|')

                    # Optionally, clean up each slide's content
                    ideas_list = [idea.strip() for idea in ideas_list]

                    if ideas_list and ideas_list[0] == '':
                        ideas_list = ideas_list[1:]

                    if not ideas_list:
                        ideas_list = "No ideas generated"  # Fallback text if API call fails or returns empty

                    content_ideas = []

                    for title in ideas_list:
                        # Assuming title format is "Diapositiva X: TITLE"
                        title_text = title.split(":")[1].strip() if len(title.split(":")) > 1 else ""
                        
                        # Generate content ideas for this title
                        prompt_for_ideas=[
                            {"role":"system","content":"Actúa como un asistente experto en la generación de ideas para "\
                            +"diapositivas en PowerPoint. No des ningún saludo, introducción ni aclaración adicional. Simplemente "\
                            +"enfócate en entregar el contenido solicitado"},
                            {"role":"user","content": f"Genera al menos 3 ideas de contenido para una diapositiva sobre: {title_text}"}
                        ]
                        
                        response_for_ideas = openai.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=prompt_for_ideas,
                            max_tokens=1000,
                            temperature=0.7,
                            n=1,  # Generate one set of ideas
                            stop=None  # Define any stop sequences if necessary
                        )
                        
                        content_ideas.append({
                            "title": title_text,
                            "ideas": response_for_ideas.choices[0].message.content.strip().split('\n')  # Correct attribute access
                        })

                        # Combine ideas_list and content_ideas for the VBA generation prompt
                        ideas_summary = " | ".join(ideas_list)
                        content_summary = " | ".join([f"{content['title']}: {'; '.join(content['ideas'])}" for content in content_ideas])

                        # Construct the conversation for VBA code generation
                        vba_conversation=[
                            {"role":"system","content":"Eres un asistente útil, especialista en programación de presentación en PowerPoint con VBA."},
                            {"role":"user","content":f"Crea un código VBA que pueda copiar y pegar en PowerPoint Visual Basic. Dicho código debe permitirme generar "\
                            f"una presentación de {ppt_slides} diapositivas sobre el tema {ppt_topic}, incluyendo los puntos clave "\
                            f"presentes en {ideas_summary}. Todas y cada una de las diapositivas deben incluir al menos un punto clave. Adicionalmente, necesito "\
                            f"que incluyas notas al pie de página con información adicional complementaria referente a {ppt_topic} que no se haya "\
                            "mencionado ya en los puntos clave. Está totalmente prohibido que las notas no aporten valor. "\
                            "Las notas deben ser únicas para diapositiva. Está prohibido también que uses una función loop del estilo "\
                            "'For i = 1 To pptPresentation.Slides.Count'. Las notas de página deben ser útiles e ir de la mano con los puntos clave. "\
                            "Comienza tu respuesta con este esquema: 'Aquí tienes el código VBA' y procede inmediatamente a mostrar el código"}
                        ]

                        # Make the OpenAI API call for VBA code generation
                        response_vba_code = openai.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=vba_conversation,
                            max_tokens=1500,  # Adjust based on your needs
                            temperature=0.7  # Adjust for creativity as needed
                        )

                        # Extract the VBA code from the response
                        vba_code = response_vba_code.choices[0].message.content

                        # You may want to store the VBA code in the MongoDB entry for the presentation
                        app.db.entries.update_one(
                            {"_id": ObjectId(entry_id)},
                            {"$set": {"vba_code": vba_code}}
                        )

                    formatted_date = datetime.datetime.today().strftime("%Y-%m-%d")
                    entry_id = app.db.entries.insert_one(
                        {"ppt_topic": ppt_topic,
                         "ppt_slides": ppt_slides,
                         "ideas_list": ideas_list,
                         "content_ideas": content_ideas,
                         "date": formatted_date,
                         "vba_code": vba_code
                        }
                    ).inserted_id

            except ValueError:
                pass

        entries_with_date = [
            (
                entry.get("ppt_topic", "No topic"),  
                entry.get("ppt_slides", "No slides"),  
                entry["date"],
                datetime.datetime.strptime(entry["date"], "%Y-%m-%d").strftime("%b %d"),
                (' '.join(entry.get("ideas_list", []))[:70]),
                str(entry["_id"]),  # Convert ObjectId to string
                ', '.join([idea['title'] for idea in entry.get("content_ideas", [])])  # Join titles of content ideas
            )
            for entry in app.db.entries.find({})
        ]

        return render_template('openai.html', entries=entries_with_date, ideas_list=ideas_list)
    
    @app.route('/entry/<entry_id>')  # New route for showing entry details
    def show_entry(entry_id):
        entry = app.db.entries.find_one({'_id': ObjectId(entry_id)})
        if entry:
            return render_template('entry_detail.html', entry=entry)
        else:
            return 'Entry not found', 404
    
    return app

# Ensure the Flask application runs with debug turned off in production
if __name__ == '__main__':
    app = create_app()
    app.run(debug=True)  # Turn off debug when deploying to production