# Electronic Assessment Form Dashboard

## Overview

The Electronic Assessment Form Dashboard is a cutting-edge web application designed to significantly streamline and enhance the academic assessment process at London South Bank University. Built using state-of-the-art Shiny for Python technology, this professional-grade platform empowers faculty members to conduct efficient, precise evaluations, generate polished and consistent PDF reports, and securely manage comprehensive assessment records through an integrated database system.

## Core Features

### Automated Student Record Integration

* Seamlessly integrates student records directly from Excel files, drastically reducing manual data entry errors.
* Essential student information includes: Student\_ID, Name, Surname, Course, Mode, Module, Title, Supervisor.

### Comprehensive Multi-Criteria Assessment

Evaluate students based on seven meticulously defined academic criteria with fully customizable weighting:

* **Research (5%)**
* **Subject Knowledge (20%)**
* **Critical Analysis (25%)**
* **Problem-Solving (30%)**
* **Practical Competence (5%)**
* **Communication (10%)**
* **Academic Integrity (5%)**

### Interactive and User-Friendly Grading Interface

* Modern, intuitive UI featuring interactive sliders, selectors, and instant grading feedback.
* Real-time score visualization for rapid and accurate grading.

### Real-Time Dynamic Grade Calculations

* Automatic weighted score computation aligned precisely with institutional assessment standards.
* Ensures consistency and fairness in grading practices across the university.

### Professional PDF Report Generation

* Generate branded, professional-quality PDF assessment reports instantly.
* Reports adhere strictly to London South Bank University's branding guidelines, ensuring uniformity and professionalism.

### Centralized and Secure Database Management

* Robust database securely storing all assessment data and feedback in structured Excel formats.
* Automatic synchronization of assessment records, guaranteeing data integrity, ease of access, and reliability.

### Advanced Academic Feedback System

* Mandates detailed, constructive feedback for exceptional grades (A+/A and F), fostering academic excellence and transparency.

## Installation

### Prerequisites

* Python 3.x
* Shiny for Python
* Pandas
* ReportLab

### Configuration Steps

1. Place the Excel file `student_records.xlsx` containing the required columns (`Student_ID, Name, Surname, Course, Mode, Module, Title, Supervisor`) in the project directory.
2. Ensure the university logo (`lsbu_logo.png`) is available in the project directory or assets folder.

### Launching the Application

Run the following command to start the dashboard:

```bash
python final_code.py
```

Access the dashboard at:

```
http://localhost:8051
```

## Assessment Workflow

1. Select the student from the interactive dropdown.
2. Grade each criterion through intuitive and interactive tools.
3. Provide comprehensive feedback, particularly for top-performing or underperforming assessments.
4. Instantly generate professional PDF assessment reports.
5. Automatically synchronize all grading information and comments with the centralized database.

## Technologies Utilized

### Backend Development

* **Python**: Primary language used for application logic, data processing, and PDF generation.

  * Libraries: Shiny, Pandas, ReportLab

### Frontend Development

* **JavaScript/jQuery**: Enhances frontend interactivity, managing dynamic UI behaviors, event handling (dropdowns, form validation), and DOM manipulation (read-only fields, styling).
* **HTML**: Structures UI components through Shiny's UI framework and custom elements (e.g., ui.tags.h1, ui.div, ui.HTML).
* **CSS**: Styles the interface, including custom styling for cards, buttons, grade selectors, responsive design, and grade-specific color themes.

## Visual Examples

* **Student Information Panel**
* **Assessment Criteria Interface**
* **Generated PDF Report Example**

## Development Team

* **M Arifur Rahman Prince**, Division of Electrical and Electronic Engineering, London South Bank University
* **Supervised by Prof. Dr. Oswaldo Cadenas**, School of Computer Science and Digital Technologies, London South Bank University

## License

This dashboard is proprietary software developed exclusively for London South Bank University faculty members. Any unauthorized use, modification, distribution, or duplication is strictly prohibited.

© 2025 London South Bank University. All rights reserved.
