from flask import Flask, request, send_file, render_template_string
from docx import Document
import os

app = Flask(__name__)
form_html = '''<!DOCTYPE html>
<html>
<head>
    <title>Admission Form</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f0f4ff; padding: 20px; }
        h2 { color: #1f4e79; text-align: center; }
        form {
            background: white; border-radius: 10px; padding: 20px;
            max-width: 700px; margin: auto;
            box-shadow: 0 0 15px rgba(0,0,0,0.1);
        }
        input, select, textarea {
            width: 100%; padding: 8px; margin: 10px 0;
            border: 1px solid #ccc; border-radius: 5px;
        }
        button {
            background: #1f4e79; color: white; padding: 10px 15px;
            border: none; border-radius: 5px; cursor: pointer;
        }
        button:hover { background: #163a5f; }
        label { font-weight: bold; display: block; margin-top: 10px; }
    </style>
    <script>
        function validateForm() {
            const form = document.forms[0];
            const email = form.Email.value;
            const phonePattern = /^[0-9]{10}$/;
            const emailPattern = /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/;
            const dob = new Date(form.DOB.value);
            const age = parseInt(form.Age.value);
            const today = new Date();
            const calculatedAge = today.getFullYear() - dob.getFullYear();
            const ageDiff = Math.abs(calculatedAge - age);

            if (!emailPattern.test(email)) {
                alert("Invalid email format");
                return false;
            }

            if (!phonePattern.test(form.FatherPhone.value) || !phonePattern.test(form.MotherPhone.value)) {
                alert("Phone number must be 10 digits");
                return false;
            }

            if (ageDiff > 1) {
                alert("Age does not match with Date of Birth");
                return false;
            }

            for (const element of form.elements) {
                if (element.hasAttribute("required") && !element.value.trim()) {
                    alert("All fields are required");
                    return false;
                }
            }
            return true;
        }
    </script>
</head>
<body>
    <h2>📄 Darussalam Public School Admission Form</h2>
    <form method="POST" onsubmit="return validateForm();">
        <input name="FullName" placeholder="Full Name" required>
        <input name="DOB" type="date" required>
        <input name="Age" placeholder="Age" required>
        <select name="Gender" required>
            <option value="">Select Gender</option>
            <option>Male</option>
            <option>Female</option>
        </select>
        <input name="Class" placeholder="Class Seeking Admission" required>
        <input name="MotherTongue" placeholder="Mother Tongue" required>
        <select name="Nationality" required>
            <option value="">Select Nationality</option>
            <option value="India">India</option>
            <option value="United States">United States</option>
            <option value="United Kingdom">United Kingdom</option>
            <option value="Canada">Canada</option>
            <option value="Australia">Australia</option>
            <option value="Germany">Germany</option>
            <option value="France">France</option>
            <option value="Italy">Italy</option>
            <option value="Spain">Spain</option>
            <option value="Brazil">Brazil</option>
            <option value="Mexico">Mexico</option>
            <option value="Russia">Russia</option>
            <option value="China">China</option>
            <option value="Japan">Japan</option>
            <option value="South Korea">South Korea</option>
            <option value="Singapore">Singapore</option>
            <option value="Malaysia">Malaysia</option>
            <option value="Thailand">Thailand</option>
            <option value="UAE">United Arab Emirates</option>
            <option value="Saudi Arabia">Saudi Arabia</option>
            <option value="South Africa">South Africa</option>
            <option value="Nigeria">Nigeria</option>
            <option value="Egypt">Egypt</option>
            <option value="Turkey">Turkey</option>
            <option value="Netherlands">Netherlands</option>
            <option value="Switzerland">Switzerland</option>
            <option value="Sweden">Sweden</option>
            <option value="Norway">Norway</option>
            <option value="Denmark">Denmark</option>
            <option value="New Zealand">New Zealand</option>
            <option value="Argentina">Argentina</option>
            <option value="Chile">Chile</option>
            <option value="Colombia">Colombia</option>
            <option value="Pakistan">Pakistan</option>
            <option value="Bangladesh">Bangladesh</option>
            <option value="Nepal">Nepal</option>
            <option value="Sri Lanka">Sri Lanka</option>
            <option value="Qatar">Qatar</option>
            <option value="Other">Other</option>
        </select>
        <select name="Religion" required>
            <option value="">Select Religion</option>
            <option value="Hinduism">Hinduism</option>
            <option value="Islam">Islam</option>
            <option value="Christianity">Christianity</option>
            <option value="Sikhism">Sikhism</option>
            <option value="Buddhism">Buddhism</option>
            <option value="Jainism">Jainism</option>
            <option value="Judaism">Judaism</option>
            <option value="Zoroastrianism">Zoroastrianism</option>
            <option value="Baháʼí Faith">Baháʼí Faith</option>
            <option value="Other">Other</option>
        </select>
        <input name="Caste" placeholder="Caste" required>
        <select name="Category" required>
            <option value="">Select Category</option>
            <option value="SC">SC</option>
            <option value="ST">ST</option>
            <option value="OBC">OBC</option>
            <option value="General">General</option>
        </select>
        <input name="Aadhar" placeholder="Aadhar Number" required>
        <input name="PreviousSchool" placeholder="Previous School" required>
        <input name="FatherName" placeholder="Father's Name" required>
        <input name="FatherPhone" placeholder="Father's Phone Number" required>
        <input name="FatherEduOcc" placeholder="Father's Qualification & Occupation" required>
        <input name="MotherName" placeholder="Mother's Name" required>
        <input name="MotherPhone" placeholder="Mother's Phone Number" required>
        <input name="MotherEduOcc" placeholder="Mother's Qualification & Occupation" required>
        <textarea name="Address" placeholder="Full Address with PIN" required rows="3"></textarea>
        <input name="Email" type="email" placeholder="Email ID" required>
        <input name="PhysicalChallenges" placeholder="Physical Challenges (if any)" required>
        <input name="SiblingDetails" placeholder="Sibling Details (if any)" required>
        <select name="Transport" required>
            <option value="">Need transportation?</option>
            <option>Yes</option>
            <option>No</option>
        </select>
        <div style="text-align: right; clear: both; margin-top: 10px;">
            <button type="submit" style="margin: 0;">Submit</button>
        </div>
    </form>
</body>
</html>
'''

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        form_data = {key: request.form[key] for key in request.form}
        docx_path = generate_filled_docx(form_data)
        return send_file(docx_path, as_attachment=True)

    return render_template_string(form_html)

from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_row_height(row, height_pt):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_pt * 20)))  # height in twips
    trHeight.set(qn('w:hRule'), 'exact')
    trPr.append(trHeight)
    
def generate_filled_docx(data):
    template_path = "static/template1.docx"
    doc = Document(template_path)

    doc.add_paragraph("\n")  # Add space

    # Table with headers
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Field'
    hdr_cells[1].text = 'Value'

    # Add data rows
    for key, value in data.items():
        row = table.add_row()
        set_row_height(row, 20)  # set row height to 20pt
        cells = row.cells
        cells[0].text = key
        cells[1].text = value

    # Pad to bottom
    min_rows = 30
    # for _ in range(min_rows - len(data)):
    #     row = table.add_row()
    #     set_row_height(row, 20)
    #     cells = row.cells
    #     cells[0].text = ""
    #     cells[1].text = ""

    output_path = "generated_admission_form.docx"
    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    app.run(debug=True)
