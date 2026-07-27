# ==================================================
# Pin Description Generator
# ==================================================

class PinDescriptionGenerator:

    @staticmethod
    def build(
            clamp,
            function,
            utilization):

        clamp = str(clamp).strip()
        function = str(function).strip()
        utilization = str(utilization).strip()

        if clamp and utilization:
            return (
                f"{clamp}"
                f"#{function}"
                f"#{utilization}"
            )
        elif clamp:
            return (
                f"{clamp}"
                f"#{function}"
            )
        elif utilization:
            return (
                f"#{function}"
                f"#{utilization}"
            )

        return function      