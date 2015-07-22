# addressSieve
Separates randomly formatted addresses into their components. Works on international addresses and zip codes.

Relies on the punctuation and relative placement of each character to determine which components are which. Uses the parserator library/module for fine adjustment of the address segments. Uses NLP (Natural Language Processing) in conjunction with Conditional Random Fields (probablistic framework) to figure out what goes where. Offers advantages over Markov models such as the hidden markov model because of its conditional nature.
