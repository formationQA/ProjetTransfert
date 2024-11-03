from scr.GuiTraitement import ImputationInterface
from scr.TraitementDonnees import ImputationProcessor

if __name__ == '__main__':
    processor = ImputationProcessor()
    interface = ImputationInterface(processor)
    interface.lancer_interface()
