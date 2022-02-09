

class Meta2DWindow():
    """
    Meta2DWindow [summary]

    [extended_summary]
    """
    def __init__(self,name,obj) -> None:
        self.name = name
        self.meta_obj = obj
        self.plot = None

    class Plot():
        """
        Plot [summary]

        [extended_summary]
        """
        def __init__(self,id,obj) -> None:
            self.id = id
            self.meta_obj = obj
            self.curve = None

        class Curve():
            """
            Curve [summary]

            [extended_summary]
            """
            def __init__(self,id,name,obj) -> None:
                self.id = id
                self.name = name
                self.meta_obj = obj
