<!-- Improved compatibility of back to top link: See: https://github.com/othneildrew/Best-README-Template/pull/73 -->
<a name="readme-top"></a>
<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Don't forget to give the project a star!
*** Thanks again! Now go create something AMAZING! :D
-->



<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->
[![Contributors][contributors-shield]]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]



<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/MauroCesp">
    <img src="images/logo.PNG" alt="Logo" width="180" height="100">
  </a>

  <h3 align="center">Reports tool</h3>

  <p align="center">
    This tools has been create to automate the process of running Webstats.
    <br />
  </p>
</div>



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

[![logo][logo]](https://example.com)

The Webstats Usage Stats Tool is a desktop application built with Python and Tkinter, designed to streamline the process of generating usage statistics reports from web statistics data. It can efficiently handle data in the form of zip files or uploaded Excel files, cleaning, formatting, and organizing the data using Python libraries like xlwings, pywin32, openpyxl, Pandas, and Numpy. 
This tool is ideal for individual reports as well as reports for consortias, significantly reducing manual work. It runs within a virtual environment to isolate it from the operating system and requires Python and Excel installations on the local machine. Users can launch the tool by creating a desktop icon pointing to the app.py file after installing the required packages listed in requirements.txt.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



### Built With

This sections has a list the main technologies, libraries, and tools used to develop the software. For example:

* ![pandas]
* ![numpy]
* ![openpyxl]
* ![python]
* ![pywin32]
* ![tkinter]


<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- GETTING STARTED -->
## Getting Started

Getting started with the Webstats Usage Stats Tool is a straightforward process. Follow these steps to begin using the tool:

### Prerequisites

Before you start using the Webstats Usage Stats Tool, make sure you have the following prerequisites in place on your computer:

- Python: Ensure that Python is installed on your local machine. You can download Python from the official website: Python Download.
- Microsoft Excel: The tool relies on Microsoft Excel for certain operations. Make sure you have Excel installed.

### Installation

To install and set up the Webstats Usage Stats Tool, follow these steps:

1. Clone the Repository: Start by cloning this repository to your local machine. You can do this by running the following command in your terminal:
   ```sh
   git clone https://github.com/MauroCesp/reports_tool.git
   ```
2. Navigate to the project directory.
   ```sh
   cd code
   ```
3. Create a virtual environment (recommended but optional): While not mandatory, it's recommended to create a virtual environment to isolate the tool's dependencies. You can create a virtual environment using the following commands:

- On windows:
   ```sh
   python -m venv venv
   venv\Scripts\activate
   ```
- On Linux and macOS
   ```sh
   python3 -m venv venv
    source venv/bin/activate
   ```   
4. Install the required packages using pip:  With the virtual environment activated (if used), install the required Python packages listed in requirements.txt:
   ```sh
   pip install -r requirements.txt
   ```


<!-- USAGE EXAMPLES -->
## Usage

- Once you've completed the installation, you're ready to start using the Webstats Usage Stats Tool. Run the tool by executing the following command:
   ```sh
   python app.py
   ```

The tool will launch, and you can specify input options and parameters as needed to generate usage statistics reports.

That's it! You're now all set to use the Webstats Usage Stats Tool to simplify the process of generating usage reports from web statistics data.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ROADMAP -->
## Roadmap

- [x] Add Changelog
- [x] Add back to top links
- [ ] Add Additional Templates w/ Examples
- [ ] Add "components" document to easily copy & paste sections of the readme
- [ ] Multi-language Support
    - [ ] Chinese
    - [ ] Spanish


<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTACT -->
## Contact

Mauro Céspedes Araya -  mauro.cespedesaraya@wolterskluwer.com

Project Link: [https://github.com/MauroCesp/reports_tool/tree/master](https://github.com/MauroCesp/reports_tool/tree/master)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ACKNOWLEDGMENTS 
## Acknowledgments

Use this space to list resources you find helpful and would like to give credit to. I've included a few of my favorites to kick things off!

* [Choose an Open Source License](https://choosealicense.com)
* [GitHub Emoji Cheat Sheet](https://www.webpagefx.com/tools/emoji-cheat-sheet)
* [Malven's Flexbox Cheatsheet](https://flexbox.malven.co/)
* [Malven's Grid Cheatsheet](https://grid.malven.co/)
* [Img Shields](https://shields.io)
* [GitHub Pages](https://pages.github.com)
* [Font Awesome](https://fontawesome.com)
* [React Icons](https://react-icons.github.io/react-icons/search)

<p align="right">(<a href="#readme-top">back to top</a>)</p>
-->



<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/othneildrew/Best-README-Template.svg?style=for-the-badge
[contributors-url]: https://github.com/othneildrew/Best-README-Template/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/othneildrew/Best-README-Template.svg?style=for-the-badge
[forks-url]: https://github.com/othneildrew/Best-README-Template/network/members
[stars-shield]: https://img.shields.io/github/stars/othneildrew/Best-README-Template.svg?style=for-the-badge
[stars-url]: https://github.com/othneildrew/Best-README-Template/stargazers
[issues-shield]: https://img.shields.io/github/issues/othneildrew/Best-README-Template.svg?style=for-the-badge
[issues-url]: https://github.com/othneildrew/Best-README-Template/issues
[license-shield]: https://img.shields.io/github/license/othneildrew/Best-README-Template.svg?style=for-the-badge
[license-url]: https://github.com/othneildrew/Best-README-Template/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://www.linkedin.com/in/mauro-cespedes-araya/
[product-screenshot]: images/screenshot.png

[numpy]: https://github.com/MauroCesp/reports_tool/blob/master/images/numpy.png
[openpyxl]: https://github.com/MauroCesp/reports_tool/blob/master/images/openpyxl.png
[pandas]: https://github.com/MauroCesp/reports_tool/blob/master/images/pandas.png
[python]: https://github.com/MauroCesp/reports_tool/blob/master/images/phyton.png
[pywin32]: https://github.com/MauroCesp/reports_tool/blob/master/images/pywin32.png
[tkinter]: https://github.com/MauroCesp/reports_tool/blob/master/images/tkinter.png
[xwings]: https://github.com/MauroCesp/reports_tool/blob/master/images/xwings.png
[Angular-url]: https://angular.io/
[Svelte.dev]: https://img.shields.io/badge/Svelte-4A4A55?style=for-the-badge&logo=svelte&logoColor=FF3E00
[Svelte-url]: https://svelte.dev/
[Laravel.com]: https://img.shields.io/badge/Laravel-FF2D20?style=for-the-badge&logo=laravel&logoColor=white
[Laravel-url]: https://laravel.com
[Bootstrap.com]: https://img.shields.io/badge/Bootstrap-563D7C?style=for-the-badge&logo=bootstrap&logoColor=white
[Bootstrap-url]: https://getbootstrap.com
[JQuery.com]: https://img.shields.io/badge/jQuery-0769AD?style=for-the-badge&logo=jquery&logoColor=white
[JQuery-url]: https://jquery.com 
[logo]:https://github.com/MauroCesp/reports_tool/blob/master/images/logo.PNG
