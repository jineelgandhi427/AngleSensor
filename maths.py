from math import pi, sin, cos, degrees, atan2


class Formula:
    def __init__(self, o_x_m, o_y_m, a_x_m, a_y_m, phi_x_m, phi_y_m, phi_m):
        self.O_X_M = o_x_m
        self.O_Y_M = o_y_m
        self.A_X_M = a_x_m
        self.A_Y_M = a_y_m
        self.PHI_X_M = phi_x_m
        self.PHI_Y_M = phi_y_m
        self.PHI_M = self.PHI_X_M - self.PHI_Y_M

    def y_sin_diff(self, sin_p, sin_n):
        y_sin_diff = sin_p - sin_n
        return y_sin_diff

    def x_cos_diff(self, cos_p, cos_n):
        x_cos_diff = cos_p - cos_n
        return x_cos_diff

    def encoder_theta_degree(self, theta_bits, prev_theta_bits):
        theta_degree = ((360/pi)*theta_bits)+prev_theta_bits
        return theta_degree

    def y1_sin(self, y_sin_diff):
        y1_sin = y_sin_diff - self.O_Y_M
        return y1_sin

    def x1_cos(self, x_cos_diff):
        x1_cos = x_cos_diff - self.O_X_M
        return x1_cos

    def y2_sin(self, y1_sin):
        y2_sin = y1_sin/self.A_Y_M
        return y2_sin

    def x2_cos(self, x1_cos):
        x2_cos = x1_cos/self.A_X_M
        return x2_cos

    def y3_sin(self, x2_cos, y2_sin):
        angle = -(self.PHI_M * pi)/180
        y3_sin = (y2_sin - x2_cos * sin(angle)) / cos(angle)
        return y3_sin

    def alpha(self, step, dir, x2_cos, y3_sin):
        if step == 0:
            angle = degrees(atan2(x2_cos, y3_sin)) - self.PHI_X_M
            angle = (angle + 360) % 360
            if dir == 'CW':
                if angle > 5:
                    return angle - 360
                else:
                    return angle
            if dir == 'CCW':
                if angle < 355:
                    return angle + 360
                else:
                    return angle
