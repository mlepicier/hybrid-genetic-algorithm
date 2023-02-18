# Hybrid Genetic Algorithm

![readme-banner](https://user-images.githubusercontent.com/108199052/206319246-08dd0a88-ecc3-432e-b0f5-5caefa8b7fa4.jpg)

## Abstract:
<p align="justify">This project consists of implementing a genetic algorithm to optimize the routing of truck deliveries to minimize transportation cost. A genetic algorithm (<a  target="_blank" href="https://en.wikipedia.org/wiki/Genetic_algorithm">GA</a>) is a metaheuristic inspired by Darwin's theory of natural selection, part of the larger class of evolutionary algorithms. With a wide range of <b>different applications</b>, beyond their common use in the field of operations research, GAs are particularly used in <b>machine learning</b> for decision tree pruning, hyper-parameter optimization (<a  target="_blank" href="https://en.wikipedia.org/wiki/Hyperparameter_optimization">HPO</a>), automated machine learning (<a  target="_blank" href="https://en.wikipedia.org/wiki/Automated_machine_learning">AutoML</a>), and neural architecture search (<a  target="_blank" href="https://en.wikipedia.org/wiki/Neural_architecture_search">NAS</a>). This project was conducted from September 2020 to January 2021 as a student at the <b>University of Technology of Troyes</b>. Below are a few paragraphs to give you context about the project.</p>

## Repository Assets:

- [Genetic Algorithm & 2-Opt Local Search](genetic-algorithm_models/tsptw) - [Full Code](genetic-algorithm_models/tsptw/ga-two-opt-tsptw.bas)
- [Genetic Algorithm & Split](genetic-algorithm_models/vrp) - [Full Code](genetic-algorithm_models/vrp/ga-split-vrp.bas)
- [TSPTW Linear Model (AMPL)](linear_model_ampl/) - [Full Code](linear_model_ampl/linear_model_tsptw.mod)
- [Project Report](genetic-algorithm_report.pdf) - To be translated to English

## Business Context:

<p align="justify">Due to the awareness of the environmental impact of individual transportation, there is an increase in the use of public transportation in the world. According to a study, in 2019, 73% of French people use buses, subways, or streetcars daily, against 63% in 2014. Urban mobility is becoming a crucial issue for most cities and will be even more so in the years to come. It is essential to control and optimize urban logistics. This problem can be applied to make decisions concerning this subject, such as school bus routes or shared-transportation design. It is also very applicable to the so-called last mile logistics and courier transport, such as the delivery of small shops.</p>


## Problem Description:

<p align="justify"> The objective of this project is to model and solve a <i><a  target="_blank" href="https://en.wikipedia.org/wiki/Travelling_salesman_problem">TSPTW</a></i> (Travelling Salesman Problem with Time Windows). In simpler words, optimize the route to visit a set of N customers and return to an initial location, equivalent to determining the optimal order of visits that minimizes the total distance traveled. Time windows in this project are flexible. That is, they can be overruled but will lead to a lateness cost penalty. The problem will be modeled in <i><a  target="_blank" href="https://en.wikipedia.org/wiki/AMPL">AMPL</a></i> and solved using a <i>Genetic Algorithm</i> including the commonly used <i><a  target="_blank" href="https://en.wikipedia.org/wiki/2-opt">2-Opt</a></i> local search move, which given the time windows constitutes the biggest algorithmic challenge of this project.</p>

<p align="justify"> As the demand grows, the vehicle capacity requires creating multiple routes, and making additional decisions around which customers to affect on each route, and still, in which order should each route visit their affected customers. This is known as the <i><a  target="_blank" href="https://en.wikipedia.org/wiki/Vehicle_routing_problem">VRP</a></i> (Vehicle Routing Problem). In 2004, Pr. Christian PRINS, Full Professor at the <b>University of Technology of Troyes</b> published a route-first cluster-second heuristic using a Hybrid Genetic Algorithm approach to solve the VRP by using the <i>Split algorithm</i> as the <i>Fitness Function</i> for the genetic algorithm. The Split optimally partitions a giant tour solution (without occurrences of the depot) into separate routes at minimum cost. This approach was implemented in this project to solve the VRP on the same dataset.</p>

***

<i>Should you have any questions, feel free to write me an [email](mailto:mlepicier.msc2022@ivey.ca), I am always happy to help.</i>
