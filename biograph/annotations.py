
class AnnotationManager:
    def __init__(self):
        self.annotations = []

    def add_annotation(self, x, y, text):
        self.annotations.append({'x': x, 'y': y, 'text': text})

    def clear(self):
        self.annotations.clear()

    def draw(self, ax, fontsize=10):
        for ann in self.annotations:
            ax.annotate(
                ann['text'],
                xy=(ann['x'], ann['y']),
                xytext=(5, 5),
                textcoords='offset points',
                fontsize=fontsize,
                color='green',
                arrowprops=dict(arrowstyle='->', color='green')
            )
